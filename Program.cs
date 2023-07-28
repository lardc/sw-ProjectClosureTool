using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.Json;
using ProjectClosureToolV2;
using Microsoft.Extensions.Configuration;
using Newtonsoft.Json;

namespace ProjectClosureToolV2
{
    public class Program
    {
        private static readonly byte[] s_nameUtf8 = Encoding.UTF8.GetBytes("name");
        private static readonly byte[] s_UrlUtf8 = Encoding.UTF8.GetBytes("shortUrl");
        private static readonly byte[] s_cardRoleUtf8 = Encoding.UTF8.GetBytes("cardRole");

        private static bool labelsListFilled;
        private static bool ignoredLabelsListFilled;

        private static int iInputIgnore;
        private static int iDeleteIgnore;
        private static bool isIgnored;

        public static List<TrelloObject> cards = new List<TrelloObject>();
        static List<TrelloObjectLabels> labels = new List<TrelloObjectLabels>();
        static List<TrelloObjectLabels> labelsList = new List<TrelloObjectLabels>();
        static List<TrelloObjectLabels> ignoredLabelsList = new List<TrelloObjectLabels>();
        static List<string> combinationsList = new List<string>();
        static List<string> combinationsListI = new List<string>();
        static IEnumerable<string> distinctCombinations = new List<string>();
        private static List<string> distinctCombinationsList;
        static IEnumerable<string> distinctCombinationsI = new List<string>();
        public static List<string> distinctCombinationsListI;

        static List<TrelloObjectLabels> cardLabels = new List<TrelloObjectLabels>();
        public static List<string> units = new List<string>();
        static IEnumerable<string> distinctUnits = new List<string>();
        public static List<string> distinctUnitsList;

        public static double sumEstimate = 0;
        public static double sumPoint = 0;

        public static List<TrelloObjectSums> sums = new List<TrelloObjectSums>();
        public static bool sumsListFilled = false;

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
            Console.WriteLine("sum - суммарные оценки (два типа) для конкретного блока для конкретной комбинации ярлыков (игнорируемые ярлыки не учитываются)");
            Console.WriteLine("sumUC - суммарные оценки (два типа) для всех блоков для всех комбинаций ярлыков (игнорируемые ярлыки не учитываются)");
            Console.WriteLine("fillExcel - формирование excel-файла");
            Console.WriteLine("configW - запись данных в конфигурационный файл");
            Console.WriteLine("configR - чтение конфигурационного файла");
            Console.WriteLine("cardsOut - вывод данных по всем карточкам");
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
        public static void ReadBoard()
        {
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
                        else if (CheckIgnored(aLabel.CardLabel))
                            Console.WriteLine($"Ярлык с номером {iInputIgnore} уже содержится в списке игнорируемых");
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
            //TableResp.iAll = 0;
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
                distinctCombinationsList = distinctCombinations.ToList();
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
            sumEstimate = 0;
            sumPoint = 0;
            string unit = distinctUnitsList.ElementAt(i - 1);
            string combination = distinctCombinationsListI.ElementAt(j - 1);
            int sumT = 0;
            foreach (TrelloObject aCard in cards)
                if (aCard.CardUnit.Equals(unit) && aCard.LabelCombinationI.Equals(combination))
                {
                    sumEstimate += aCard.CardEstimate;
                    sumPoint += aCard.CardPoint;
                    sumT++;
                    Console.WriteLine();
                }
            if (sumT > 0) 
            {
                Console.WriteLine($"Блок: {unit}");
                Console.WriteLine($"Комбинация ярлыков: {combination}");
                Console.WriteLine($"Суммарное оценочное значение: ({sumEstimate}).\nСуммарное реальное значение: [{sumPoint}].");
                Console.WriteLine();
            }
        }

        public static void WriteConfig(string rr)
        {
            var obj = new ConfProg()
            {
                APIKey = API_Req.APIKey,
                myTrelloToken = API_Req.myTrelloToken,
                boardCode = API_Req.boardCode,
                IgnoredLabels = ignoredLabelsList
            };
            var json = JsonConvert.SerializeObject(obj);
            if (!File.Exists(rr)) { using var stream = File.Create(rr, (int)FileMode.Create) ; }
            File.WriteAllText($"{rr}", json);
        }

        public static void ReadConfig(string rr)
        {
            ClearIgnoredLabels();
            var json = File.ReadAllText($"{rr}");
            var obj = JsonConvert.DeserializeObject<ConfProg>(json);
            API_Req.APIKey = obj.APIKey;
            API_Req.myTrelloToken = obj.myTrelloToken;
            API_Req.boardCode = obj.boardCode;
            Console.WriteLine($"APIKey = {API_Req.APIKey}");
            Console.WriteLine($"myTrelloToken = {API_Req.myTrelloToken}");
            ReadBoard();
            BoardM();
            foreach (TrelloObjectLabels aLabel in obj.IgnoredLabels)
            {
                if (labels.Contains(aLabel))
                {
                    ignoredLabelsList.Add(aLabel);
                    ignoredLabelsListFilled = true;
                }
            }
            if (ignoredLabelsListFilled)
                Console.WriteLine("Игнорируемые ярлыки:");
            IgnoredLabelsList();
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
                        labelsListFilled = false;
                        Console.Write("Введите код доски >");
                        try
                        { API_Req.boardCode = Console.ReadLine(); }
                        catch (Exception e)
                        {
                            Console.WriteLine(e.Message);
                            Console.WriteLine("Press any key");
                            Console.ReadKey();
                            return;
                        }
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
                            Console.WriteLine($"{distinctCombinationsListI.IndexOf(aCombination) + 1}.\n{aCombination}");
                        break;
                    case "sum":
                        sumEstimate = 0;
                        sumPoint = 0;
                        UnitsList();
                        Console.WriteLine();
                        LabelCombinationsI();
                        foreach (string aCombination in distinctCombinationsListI)
                            Console.WriteLine($"{distinctCombinationsListI.IndexOf(aCombination) + 1}. {aCombination}");
                        Console.WriteLine();
                        Console.Write("Введите номер блока >");
                        int iUnit = int.Parse(Console.ReadLine());
                        if (iUnit > distinctUnitsList.Count) { Console.WriteLine("Введите корректный номер блока"); }
                        else
                        {
                            Console.Write("Введите номер комбинации >");
                            int iCombination = int.Parse(Console.ReadLine());
                            if (iCombination > distinctCombinationsListI.Count) { Console.WriteLine("Введите корректный номер комбинации"); }
                            else Sum(iUnit, iCombination);
                        }
                        break;
                    case "sums":
                        LabelCombinationsI();
                        Console.WriteLine();
                        UnitsList();
                        Console.WriteLine();
                        Console.Write("Введите номер блока >");
                        iUnit = int.Parse(Console.ReadLine());
                        if (iUnit > distinctUnitsList.Count) { Console.WriteLine("Введите корректный номер блока"); }
                        else for (int i = 0; i < distinctCombinationsI.ToList().Count; i++)
                            Sum(iUnit, i + 1);
                        break;
                    case "sumUC":
                        int iU = 0;
                        int iC = 0;
                        UnitsList();
                        LabelCombinationsI();
                        for (iU = 1; iU <= distinctUnitsList.Count; iU++)
                        {
                            for (iC = 1; iC <= distinctCombinationsListI.Count; iC++)
                            {
                                Console.WriteLine($"Блок {iU}.  Комбинация {iC}.");
                                Sum(iU, iC);
                                sums.Add(new TrelloObjectSums()
                                {
                                    CardUnit = distinctUnitsList.ElementAt(iU - 1),
                                    LabelCombinationI = distinctCombinationsListI.ElementAt(iC - 1),
                                    SumEstimate = sumEstimate,
                                    SumPoint = sumPoint
                                });
                            }
                        }
                        sumsListFilled = true;
                        foreach (TrelloObjectSums aSum in sums)
                        {
                            Console.WriteLine(aSum.ToString());
                        }
                        break;
                    case "configW":
                        string fileName = "config.json";
                        try { WriteConfig(fileName); }
                        catch (Exception e)
                        {
                            Console.WriteLine(e.Message);
                            Console.WriteLine("Press any key");
                            Console.ReadKey();
                            return;
                        }
                        break;
                    case "configR":
                        fileName = "config.json";
                        if (File.Exists(fileName))
                        {
                            try { ReadConfig(fileName); }
                            catch (Exception e)
                            {
                                Console.WriteLine(e.Message);
                                Console.WriteLine("Press any key");
                                Console.ReadKey();
                                return;
                            }
                        }
                        else { Console.WriteLine($"Нет конфигурационного файла"); }
                        break;
                    case "fillExcel":
                        if (!labelsListFilled) { Console.WriteLine("Код доски не введён"); break; }
                        if (sumsListFilled)
                        {
                            Console.WriteLine("Формирование excel-файла");
                            TableResp.FillExcel();
                        }
                        else { Console.WriteLine("Список сумм не сформирован. Введите команду <sumUC>"); }
                        break;
                    case "cardsOut":
                        LabelCombinationsI();
                        foreach (TrelloObject aCard in cards)
                            Console.WriteLine(aCard.ToString());
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