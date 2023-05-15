using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.Json;
using ProjectClosureToolV2;

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

        private static int iInputIgnore;
        private static int iDeleteIgnore;
        private static bool isIgnored;

        static List<TrelloObject> cards = new List<TrelloObject>();
        static List<TrelloObjectLabels> labels = new List<TrelloObjectLabels>();
        static List<TrelloObjectLabels> labelsList = new List<TrelloObjectLabels>();
        static List<TrelloObjectLabels> ignoredLabelsList = new List<TrelloObjectLabels>();
        static List<string> combinationsList = new List<string>();
        static List<string> combinationsListI = new List<string>();
        static IEnumerable<string> distinctCombinations = new List<string>();
        private static List<string> distinctCombinationsList;
        static IEnumerable<string> distinctCombinationsI = new List<string>();
        private static List<string> distinctCombinationsListI;

        static List<TrelloObjectLabels> cardLabels = new List<TrelloObjectLabels>();
        public static List<string> units = new List<string>();
        static IEnumerable<string> distinctUnits = new List<string>();
        private static List<string> distinctUnitsList;

        //public static void NameToken(Utf8JsonReader reader)
        //{
        //    if (reader.GetString().StartsWith("name"))
        //    {
        //        // Чтение токена
        //        reader.Read();

        //        // Блок? 
        //        if (reader.CurrentDepth.Equals(2))
        //            // Запись оценочных и реальных значений для карточки
        //            Trl.Fill_Current_Values(reader.GetString().ToString());
        //        // Стадия? Команда?
        //        else if (reader.CurrentDepth.Equals(4))
        //            Trl.Search_Departments_Teams(reader.GetString().ToString());
        //    }
        //}

        //public static void ShortUrlToken(Utf8JsonReader reader)
        //{
        //    if (reader.GetString().StartsWith("shortUrl"))
        //    {
        //        // Чтение токена
        //        reader.Read();
        //        TableResp.currentShortUrl = reader.GetString().ToString();
        //        //Trl.cardURL[Trl.iCard] = TableResp.currentShortUrl;
        //    }
        //}

        //public static void BadgesToken(Utf8JsonReader reader)
        //{
        //    if (reader.GetString().StartsWith("badges"))
        //    {
        //        // Чтение токена
        //        reader.Read();
        //        TableResp.badges++;
        //    }
        //}

        //public static void CardRoleToken(Utf8JsonReader reader)
        //{
        //    if (reader.GetString().StartsWith("cardRole"))
        //    {
        //        // Чтение токена
        //        reader.Read();
        //        TableResp.cardRole++;
        //        //Trl.FillParsedValues();
        //    }
        //}

        public static void Help()
        {
            Console.WriteLine("kt - ввод ключа и токена");
            Console.WriteLine("board - ввод кода доски");
            Console.WriteLine("labels - нумерованный список всех имеющихся на доске ярлыков");
            Console.WriteLine("addI - создание списка игнорируемых ярлыков");
            Console.WriteLine("listI - просмотр списка игнорируемых ярлыков");
            Console.WriteLine("deleteI - удаление игнорируемых ярлыков по номеру");
            Console.WriteLine("clearI - очистка списка игнорируемых ярлыков");
            Console.WriteLine("boardM - ввод доски в память");
            Console.WriteLine("units - вывод списка всех блоков на доске");
            Console.WriteLine("labelC - список всех комбинаций ярлыков");
            Console.WriteLine("labelCI - список всех комбинаций ярлыков (игнорируемые ярлыки не учитываются)");
            Console.WriteLine("sums - суммарные оценки (два типа) для конкретного блока для всех уникальных комбинаций ярлыков (игнорируемые ярлыки не учитываются)");
            Console.WriteLine("sum - суммарные оценки (два типа) для конкретного блока для конкретной комбинации ярлыков ярлыков (игнорируемые ярлыки не учитываются)");
            //Console.WriteLine("cardsOut - вывод данных по всем карточкам");
            //Console.WriteLine("cardOut - вывод ярлыков для заданной карточки");
            //Console.WriteLine("fillExcel - формирование excel-файла");
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

        //public static void Board()
        //{
        //    Trl.ClearCurrentUnit();
        //    Trl.ClearValues();
        //    IgnoredLabelsList(nIgnore);

        //    try
        //    {
        //        ReadOnlySpan<byte> s_readToEnd_stringUtf8 = Encoding.UTF8.GetBytes(API_Req.readToEnd_string);
        //        var reader = new Utf8JsonReader(s_readToEnd_stringUtf8);
        //        Trl.ClearParsedCardErrorData();
        //        while (reader.Read())
        //        {
        //            // Тип считанного токена
        //            JsonTokenType tokenType;

        //            tokenType = reader.TokenType;
        //            switch (tokenType)
        //            {
        //                // Тип токена - начало объекта JSON
        //                case JsonTokenType.StartObject:
        //                    break;
        //                // Тип токена - название свойства
        //                case JsonTokenType.PropertyName:
        //                    // Это токен "badges"?
        //                    if (reader.ValueTextEquals(s_badgesUtf8))
        //                    {
        //                        try { BadgesToken(reader); }
        //                        catch (Exception e)
        //                        {
        //                            Console.WriteLine(e.Message);
        //                            Console.WriteLine("Press any key");
        //                            Console.ReadKey();
        //                            return;
        //                        }
        //                        if (Trl.parseBadgesToken)
        //                        {
        //                            if (Trl.iErrorParse < Trl.iMaxParse) Trl.iErrorParse++;
        //                            Trl.parseStrErrorMessage[Trl.iErrorParse] = "Нет конца карточки";
        //                            Trl.parseStrErrorCardURL[Trl.iErrorParse] = TableResp.currentShortUrl;
        //                            Trl.FillParsedValues();
        //                            Trl.parseBadgesToken = true;
        //                        }
        //                        else { Trl.parseBadgesToken = true; }
        //                    }
        //                    // Это токен "name"?
        //                    else if (reader.ValueTextEquals(s_nameUtf8))
        //                    {
        //                        try { NameToken(reader); }
        //                        catch (Exception e)
        //                        {
        //                            Console.WriteLine(e.Message);
        //                            Console.WriteLine("Press any key");
        //                            Console.ReadKey();
        //                            return;
        //                        }
        //                    }
        //                    // Это токен "shortUrl"?
        //                    else if (reader.ValueTextEquals(s_UrlUtf8))
        //                    {
        //                        try { ShortUrlToken(reader); }
        //                        catch (Exception e)
        //                        {
        //                            Console.WriteLine(e.Message);
        //                            Console.WriteLine("Press any key");
        //                            Console.ReadKey();
        //                            return;
        //                        }
        //                    }
        //                    // Это токен "cardRole"? 
        //                    else if (reader.ValueTextEquals(s_cardRoleUtf8))
        //                    {
        //                        try { CardRoleToken(reader); }
        //                        catch (Exception e)
        //                        {
        //                            Console.WriteLine(e.Message);
        //                            Console.WriteLine("Press any key");
        //                            Console.ReadKey();
        //                            return;
        //                        }
        //                        Trl.FillParsedValues();
        //                    }
        //                    break;
        //            }
        //        }
        //    }
        //    catch (Exception e)
        //    {
        //        Console.WriteLine(e.Message);
        //        Console.WriteLine("Press any key");
        //        Console.ReadKey();
        //        return;
        //    }

        //    Trl.FillParsedValues();
        //    labelsListFilled = true;
        //}

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

        //public static void ReadLabels()
        //{
        //    try 
        //    {
        //        ReadOnlySpan<byte> s_readToEnd_stringUtf8 = Encoding.UTF8.GetBytes(API_Req.readToEnd_string);
        //        var reader = new Utf8JsonReader(s_readToEnd_stringUtf8);
        //        //Trl.ClearParsedCardErrorData();
        //        ClearLabels();
        //        while (reader.Read())
        //        {
        //            JsonTokenType tokenType;

        //            tokenType = reader.TokenType;
        //            switch (tokenType)
        //            {
        //                // Тип токена - начало объекта JSON
        //                case JsonTokenType.StartObject:
        //                    {
        //                        break;
        //                    }
        //                // Тип токена - название свойства
        //                case JsonTokenType.PropertyName:
        //                    if (reader.ValueTextEquals(s_nameUtf8))
        //                    {
        //                        if (reader.GetString().StartsWith("name"))
        //                        try
        //                        {
        //                            reader.Read(); 
        //                            if (reader.CurrentDepth.Equals(4)) Trl.Input_Labels(reader.GetString().ToString());
        //                        }
        //                        catch (Exception e)
        //                        {
        //                            Console.WriteLine(e.Message);
        //                            Console.WriteLine("Press any key");
        //                            Console.ReadKey();
        //                            return;
        //                        }
        //                    }
        //                    break;
        //            }
        //        }
        //    }
        //    catch (Exception e)
        //    {
        //        Console.WriteLine(e.Message);
        //        Console.WriteLine("Press any key");
        //        Console.ReadKey();
        //        return;
        //    }
        //    labelsListFilled = true;
        //}

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
                foreach (TrelloObjectLabels aLabel in labelsList)
                    if (!CheckIgnored(aLabel.CardLabel))
                        Console.WriteLine(aLabel);
            }
        }

        public static void BoardM()
        {
            ClearCardM();
            cards.Clear();
            units.Clear();
            labels.Clear();
            combinationsList.Clear();
            distinctUnits = Enumerable.Empty<string>();
            TableResp.iAll = 0;
            try
            {
                ReadOnlySpan<byte> s_readToEnd_stringUtf8 = Encoding.UTF8.GetBytes(API_Req.readToEnd_string);
                var reader = new Utf8JsonReader(s_readToEnd_stringUtf8);
                while (reader.Read())
                {
                    JsonTokenType tokenType;
                    tokenType = reader.TokenType;
                    switch (tokenType)
                    {
                        case JsonTokenType.StartObject:
                            break;
                        case JsonTokenType.PropertyName:
                            if (reader.ValueTextEquals(s_nameUtf8))
                            {
                                try { NameTokenM(reader); }
                                catch (Exception e)
                                {
                                    Console.WriteLine(e.Message);
                                    Console.WriteLine("Press any key");
                                    Console.ReadKey();
                                    return;
                                }
                            }
                            else if (reader.ValueTextEquals(s_UrlUtf8))
                            {
                                try { ShortUrlTokenM(reader); }
                                catch (Exception e)
                                {
                                    Console.WriteLine(e.Message);
                                    Console.WriteLine("Press any key");
                                    Console.ReadKey();
                                    return;
                                }
                            }
                            else if (reader.ValueTextEquals(s_cardRoleUtf8))
                            {
                                try { CardRoleTokenM(reader); }
                                catch (Exception e)
                                {
                                    Console.WriteLine(e.Message);
                                    Console.WriteLine("Press any key");
                                    Console.ReadKey();
                                    return;
                                }
                                ClearCardM();
                            }
                            break;
                    }
                }
                labels.Sort();
                LabelCombinationsI();
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                Console.WriteLine("Press any key");
                Console.ReadKey();
                return;
            }
        }
        public static void NameTokenM(Utf8JsonReader reader)
        {
            if (reader.GetString().StartsWith("name"))
            {
                reader.Read();
                if (reader.CurrentDepth.Equals(2))
                    try { Trl.SearchUnitValuesM(reader.GetString().ToString()); }
                    catch (Exception e)
                    {
                        Console.WriteLine(e.Message);
                        Console.WriteLine("Press any key");
                        Console.ReadKey();
                        return;
                    }
                else if (reader.CurrentDepth.Equals(4))
                    try { Trl.SearchLabelsM(reader.GetString().ToString()); }
                    catch (Exception e)
                    {
                        Console.WriteLine(e.Message);
                        Console.WriteLine("Press any key");
                        Console.ReadKey();
                        return;
                    }
            }
        }

        public static void ClearCardM()
        {
            Trl.currentCardURL = "";
            Trl.currentCardUnit = "";
            Trl.currentCardName = "";
            Trl.currentCardEstimate = 0;
            Trl.currentCardPoint = 0;
            for (int i = 0; i < 20; i++)
                Trl.cardLabels[i] = "";
            Trl.iLabels = 0;
        }

        public static void ShortUrlTokenM(Utf8JsonReader reader)
        {
            if (reader.GetString().StartsWith("shortUrl"))
            {
                reader.Read();
                Trl.currentCardURL = reader.GetString().ToString();
            }
        }

        public static void CardRoleTokenM(Utf8JsonReader reader)
        {
            int newCardID = cards.Count;
            cards.Add(new TrelloObject()
            {
                CardID = newCardID,
                CardURL = Trl.currentCardURL,
                CardUnit = Trl.currentCardUnit,
                CardName = Trl.currentCardName,
                CardEstimate = Trl.currentCardEstimate,
                CardPoint = Trl.currentCardPoint
            });
            for (int i = 0; i < Trl.iLabels; i++)
            {
                labels.Add(new TrelloObjectLabels()
                {
                    CardID = newCardID,
                    CardLabel = Trl.cardLabels[i]
                });
                bool contains = false;
                foreach (TrelloObjectLabels aLabel in labelsList)
                {
                    if (aLabel.CardLabel.Equals(Trl.cardLabels[i]))
                        contains = true;
                }
                if (!contains)
                {
                    labelsList.Add(new TrelloObjectLabels()
                    {
                        CardID = labelsList.Count,
                        CardLabel = Trl.cardLabels[i]
                    });
                    labelsListFilled = true;
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
                foreach (TrelloObjectLabels aLabel in labelsList)
                    if (aLabel.CardID.Equals(iInputIgnore - 1))
                    {
                        if (!CheckLabels(aLabel.CardLabel) || iInputIgnore > labelsList.Count)
                            Console.WriteLine($"Нет ярлыка с номером <{iInputIgnore}>. Введите другой номер.");
                        else
                        {
                            if (isIgnored)
                            {
                                foreach (TrelloObjectLabels label in labelsList)
                                    if (label.CardID.Equals(iInputIgnore - 1))
                                        Console.WriteLine($"Ярлык <{label.CardLabel}> уже содержится в списке игнорируемых");
                            }
                            else
                            {
                                ignoredLabelsList.Add(aLabel);
                                ignoredLabelsListFilled = true;
                            }
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
            foreach (TrelloObject aCard in cards)
                aCard.LabelCombinationI = "";
            LabelCombinationsI();
        }

        // Просмотр списка игнорируемых ярлыков
        public static void IgnoredLabelsList()
        {
            if (!ignoredLabelsListFilled)
                Console.WriteLine("Игнорируемых ярлыков нет");
            else
                foreach (TrelloObjectLabels aLabel in ignoredLabelsList)
                {
                    if (!CheckLabels(aLabel.CardLabel))
                        Console.WriteLine($"Игнорируемый ярлык <{aLabel.CardLabel}> отсутствует в списке ярлыков. Перезагрузите доску.");
                    else Console.WriteLine(aLabel);
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
                IgnoredLabelsList();
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
                    var temp = new List<TrelloObjectLabels>();
                    bool contains = false;
                    foreach (TrelloObjectLabels aLabel in ignoredLabelsList)
                    {
                        if (aLabel.CardID.Equals(iDeleteIgnore - 1))
                        {
                            contains = true;
                            if (aLabel.CardID.Equals(iDeleteIgnore - 1))
                                temp.Add(aLabel);
                        }
                    }
                    if (!contains)
                        Console.WriteLine($"Нет игнорируемого ярлыка с номером <{iDeleteIgnore}>. Введите другой номер.");
                    if (ignoredLabelsList.Count < 1)
                    {
                        ignoredLabelsListFilled = false;
                        iDeleteIgnore = 0;
                        IgnoredLabelsList();
                    }
                    else
                    {
                        foreach (TrelloObjectLabels aLabel in temp)
                            if (ignoredLabelsList.Contains(aLabel))
                                ignoredLabelsList.Remove(aLabel);
                        IgnoredLabelsList();
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
            LabelCombinationsI();
        }

        // Очистка списка игнорируемых ярлыков
        public static void ClearIgnoredLabels()
        {
            ignoredLabelsList.Clear();
            ignoredLabelsListFilled = false;
            LabelCombinationsI();
        }

        public static void ClearLabels()
        {
            labelsList.Clear();
            Trl.iLabels = 0;
            TableResp.iAll = 0;
        }

        public static bool CheckLabels(string rr)
        {
            bool isLabel = false;
            foreach (TrelloObjectLabels aLabel in labelsList)
            {
                if (aLabel.CardLabel.Equals(rr))
                    isLabel = true;
            }
            return isLabel;
        }

        public static bool CheckIgnored(string rr)
        {
            bool isIgnored = false;
            foreach (TrelloObjectLabels aLabel in ignoredLabelsList)
            {
                if (aLabel.CardLabel.Equals(rr))
                    isIgnored = true;
            }
            return isIgnored;
        }

        public static void UnitsList()
        {
            distinctUnits = units.Distinct();
            distinctUnitsList = distinctUnits.ToList();
            distinctUnitsList.Sort();
            foreach (string aUnit in distinctUnitsList)
                Console.WriteLine($"{distinctUnitsList.IndexOf(aUnit) + 1}. {aUnit}");
        }

        public static void LabelCombinations()
        {
            try
            {
                combinationsList.Clear();
                distinctCombinations = Enumerable.Empty<string>();
                labels.Sort();
                for (int i = 0; i < cards.Count; i++)
                {
                    string sCombination = "";
                    foreach (TrelloObjectLabels aLabel in labels)
                        if (aLabel.CardID.Equals(i))
                            sCombination += $"{aLabel.CardLabel}        ";
                    foreach (TrelloObject aCard in cards)
                        if (aCard.CardID.Equals(i) && sCombination != "")
                        {
                            aCard.LabelCombination = sCombination;
                            combinationsList.Add(sCombination);
                        }
                }
                distinctCombinations = combinationsList.Distinct();
                distinctCombinationsList = distinctCombinationsI.ToList();
                distinctCombinationsList.Sort();
                foreach (string aCombination in distinctCombinationsList)
                    Console.WriteLine(aCombination);
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                Console.WriteLine("Press any key");
                Console.ReadKey();
                return;
            }
        }

        public static void LabelCombinationsI()
        {
            try
            {
                combinationsListI.Clear();
                distinctCombinationsI = Enumerable.Empty<string>();
                labels.Sort();
                for (int i = 0; i < cards.Count; i++)
                {
                    string sCombination = "";
                    foreach (TrelloObjectLabels aLabel in labels)
                        if (aLabel.CardID.Equals(i) && !ignoredLabelsList.Contains(aLabel))
                            sCombination += $"{aLabel.CardLabel}\n";
                    foreach (TrelloObject aCard in cards)
                        if (aCard.CardID.Equals(i) && sCombination != "")
                        {
                            aCard.LabelCombinationI = sCombination;
                            combinationsListI.Add(sCombination);
                        }
                }
                distinctCombinationsI = combinationsListI.Distinct();
                distinctCombinationsListI = distinctCombinationsI.ToList();
                distinctCombinationsListI.Sort();
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                Console.WriteLine("Press any key");
                Console.ReadKey();
                return;
            }
        }

        public static void CardOutput(string rr)
        {
            cardLabels.Clear();
            int id = -1;
            foreach (TrelloObject aCard in cards)
                if (aCard.CardURL.Equals(rr))
                    id = aCard.CardID;
            if (id < 0)
                Console.WriteLine("Карточка не найдена");
            else
            {
                foreach (TrelloObjectLabels aLabel in labels)
                    if (aLabel.CardID.Equals(id))
                        cardLabels.Add(aLabel);
                foreach (TrelloObjectLabels aLabel in cardLabels)
                    Console.WriteLine(aLabel);
            }
        }

        public static void Sum(int i, int j)
        {
            double sumEstimate = 0;
            double sumPoint = 0;
            string unit = distinctUnitsList.ElementAt(i - 1);
            string combination = distinctCombinationsListI.ElementAt(j - 1);
            foreach (TrelloObject aCard in cards)
                if (aCard.CardUnit.Equals(unit) && aCard.LabelCombinationI.Equals(combination))
                {
                    sumEstimate += aCard.CardEstimate;
                    sumPoint += aCard.CardPoint;
                }
            Console.WriteLine($"Блок: {unit}");
            Console.WriteLine($"Комбинация ярлыков: {combination}");
            Console.WriteLine($"Суммарное оценочное значение: ({sumEstimate}).\n" +
                $"Суммарное реальное значение: [{sumPoint}].");
            Console.WriteLine();
        }

        static void Main()
        {
            labels.Sort(delegate (TrelloObjectLabels x, TrelloObjectLabels y)
            {
                if (x.CardLabel == null && y.CardLabel == null) return 0;
                else if (x.CardLabel == null) return -1;
                else if (y.CardLabel == null) return 1;
                else return x.CardLabel.CompareTo(y.CardLabel);
            });

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
                        BoardM();
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
                        IgnoredLabelsList();
                        break;
                    case "deleteI":
                        DeleteIgnoredLabel();
                        break;
                    case "clearI":
                        ClearIgnoredLabels();
                        break;
                    case "units":
                        UnitsList();
                        break;
                    case "labelC":
                        LabelCombinations();
                        break;
                    case "labelCI":
                        LabelCombinationsI();
                        foreach (string aCombination in distinctCombinationsListI)
                            Console.WriteLine($"{distinctCombinationsListI.IndexOf(aCombination) + 1}.\n" +
                                $"{aCombination}");
                        break;
                    case "sum":
                        UnitsList();
                        Console.WriteLine();
                        LabelCombinationsI();
                        foreach (string aCombination in distinctCombinationsListI)
                            Console.WriteLine($"{distinctCombinationsListI.IndexOf(aCombination) + 1}. {aCombination}");
                        Console.WriteLine();
                        Console.Write("Введите номер блока >");
                        int iUnit = int.Parse(Console.ReadLine());
                        Console.Write("Введите номер комбинации >");
                        int iCombination = int.Parse(Console.ReadLine());
                        Sum(iUnit, iCombination);
                        break;
                    case "sums":
                        LabelCombinationsI();
                        Console.WriteLine();
                        UnitsList();
                        Console.WriteLine();
                        Console.Write("Введите номер блока >");
                        iUnit = int.Parse(Console.ReadLine());
                        for (int i = 0; i < distinctCombinationsI.ToList().Count; i++)
                            Sum(iUnit, i + 1);
                        break;

                    //case "fillExcel":
                    //    if (!labelsListFilled) Console.WriteLine("Код доски не введён");
                    //    else
                    //    {
                    //        Console.WriteLine("Формирование excel-файла");
                    //        Board();
                    //        Trl.FillExcel();
                    //    }
                    //    break;
                    case "cardsOut":
                        LabelCombinationsI();
                        foreach (TrelloObject aCard in cards)
                            Console.WriteLine(aCard.ToString());
                        break;
                    case "cardOut":
                        LabelCombinationsI();
                        Console.Write("Введите shortURL карточки >");
                        string cShortURL = Console.ReadLine();
                        CardOutput(cShortURL);
                        break;
                    case "boardM":
                        BoardM();
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