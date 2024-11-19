using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;

namespace ProjectClosureToolMVVM
{
    internal class Download
    {
        public static List<TrelloObject> cards = [];
        public static List<TrelloObjectLabels> labels = [];
        public static List<TrelloObjectLabels> labelsList = [];
        public static List<TrelloObjectLabels> ignoredLabelsList = [];
        public static List<string> units = [];
        public static IEnumerable<string> distinctUnits = [];
        public static List<string> distinctUnitsList;
        public static List<string> selectedUnits = [];
        private static readonly List<string> combinationsListI = [];
        static IEnumerable<string> distinctCombinationsI = [];
        public static List<string> distinctCombinationsListI = [];
        public static List<string> selectedCombinations = [];

        public static bool distinctCombinationsListFilled = false;
        public static bool labelsListFilled = false;
        public static bool ignoredLabelsListFilled = false;
        public static bool unitsListFilled = false;
        public static string currentCardURL;
        public static string currentCardUnit;
        public static double currentCardEstimate;
        public static double currentCardPoint;
        public static double sumEstimate;
        public static double sumPoint;
        public static string[] cardLabels = new string[1000];

        public static void ClearLabels()
        {
            labelsList.Clear();
            Trl.iLabels = 0;
        }

        public static void ClearCardM()
        {
            currentCardURL = "";
            currentCardUnit = "";
            currentCardEstimate = 0;
            currentCardPoint = 0;
            for (int i = 0; i < 20; i++)
                cardLabels[i] = "";
            Trl.iLabels = 0;
        }

        public static void NameTokenM(Utf8JsonReader reader)
        {
            if (reader.GetString().StartsWith("name"))
            {
                reader.Read();
                if (reader.CurrentDepth.Equals(2))
                    Trl.SearchUnitValuesM(reader.GetString().ToString());
                else if (reader.CurrentDepth.Equals(4))
                    Trl.SearchLabelsM(reader.GetString().ToString());
            }
        }

        public static void ShortUrlTokenM(Utf8JsonReader reader)
        {
            if (reader.GetString().StartsWith("shortUrl"))
            {
                reader.Read();
                currentCardURL = reader.GetString().ToString();
            }
        }

        public static void CardRoleTokenM(Utf8JsonReader reader)
        {
            int newCardID = cards.Count;
            cards.Add(new TrelloObject()
            {
                CardID = newCardID,
                CardURL = currentCardURL,
                CardUnit = currentCardUnit,
                CardEstimate = currentCardEstimate,
                CardPoint = currentCardPoint
            });
            if (Trl.iLabels > 0) for (int i = 0; i < Trl.iLabels; i++)
            {
                labels.Add(new TrelloObjectLabels()
                {
                    CardID = newCardID,
                    CardLabel = cardLabels[i]
                });
                bool contains = false;
                foreach (TrelloObjectLabels aLabel in labelsList)
                {
                    if (aLabel.CardLabel.Equals(cardLabels[i]))
                        contains = true;
                }
                if (!contains)
                {
                    labelsList.Add(new TrelloObjectLabels()
                    {
                        CardID = labelsList.Count,
                        CardLabel = cardLabels[i]
                    });                    
                    labelsListFilled = true;
                }
            }
        }

        public static void BoardM()
        {
            ClearCardM();
            cards.Clear();
            units.Clear();
            labels.Clear();
            combinationsListI.Clear();
            distinctUnits = [];
            ReadOnlySpan<byte> s_readToEnd_stringUtf8 = Encoding.UTF8.GetBytes(API_Req.ReadToEnd_string);
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
                        if (reader.ValueTextEquals(API_Req.s_nameUtf8))
                        {
                            NameTokenM(reader);
                        }
                        else if (reader.ValueTextEquals(API_Req.s_UrlUtf8))
                        {
                            ShortUrlTokenM(reader);
                        }
                        else if (reader.ValueTextEquals(API_Req.s_cardRoleUtf8))
                        {
                            CardRoleTokenM(reader);
                            ClearCardM();
                        }
                        break;
                }
            }
            labels.Sort();
        }

        public static bool CheckIgnored(string rr)
        {
            bool isIgnored = false;
            if (ignoredLabelsListFilled) foreach (TrelloObjectLabels aLabel in ignoredLabelsList)
            {
                if (aLabel.CardLabel.Equals(rr))
                    isIgnored = true;
            }
            return isIgnored;
        }

        public static void LabelCombinationsI()
        {
            combinationsListI.Clear();
            distinctCombinationsListI.Clear();
            distinctCombinationsI = [];
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
            distinctCombinationsListFilled = true;
        }
    }
}
