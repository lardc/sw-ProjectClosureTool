﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;

namespace ProjectClosureToolMVVM
{
    internal class Download
    {
        private static List<TrelloObject> cards = new List<TrelloObject>();
        public static List<TrelloObjectLabels> labels = new List<TrelloObjectLabels>();
        public static List<TrelloObjectLabels> labelsList = new List<TrelloObjectLabels>();
        public static List<TrelloObjectLabels> ignoredLabelsList = new List<TrelloObjectLabels>();
        public static List<string> units = new List<string>();
        public static IEnumerable<string> distinctUnits = new List<string>();
        public static List<string> distinctUnitsList;
        public static List<string> selectedUnits = new List<string>();

        public static bool labelsListFilled = false;
        //public static int iLabels;
        public static bool ignoredLabelsListFilled = false;
        public static bool unitsListFilled = false;
        public static string currentCardURL;
        public static string currentCardUnit;
        public static string currentCardName;
        public static double currentCardEstimate;
        public static double currentCardPoint;
        public static double sumEstimate;
        public static double sumPoint;
        public static string[] cardLabels = new string[1000];

        //private static bool isIgnoredDGV2 = false;
        //private static bool isCheckedDGV3 = false;
        //private static bool isCheckedDGV4 = false;

        public static void ClearLabels()
        {
            labelsList.Clear();
            Trl.iLabels = 0;
        }

        public static void ClearCardM()
        {
            currentCardURL = "";
            currentCardUnit = "";
            //currentCardName = "";
            //currentCardEstimate = 0;
            //currentCardPoint = 0;
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
                //try { Trl.SearchUnitValuesM(reader.GetString().ToString()); }
                //catch (Exception e)
                //{
                //    Console.WriteLine(e.Message);
                //    Console.WriteLine("Press any key");
                //    Console.ReadKey();
                //    return;
                //}
                else if (reader.CurrentDepth.Equals(4))
                    Trl.SearchLabelsM(reader.GetString().ToString());
                //try { Trl.SearchLabelsM(reader.GetString().ToString()); }
                //catch (Exception e)
                //{
                //    Console.WriteLine(e.Message);
                //    Console.WriteLine("Press any key");
                //    Console.ReadKey();
                //    return;
                //}
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
                CardName = currentCardName,
                CardEstimate = currentCardEstimate,
                CardPoint = currentCardPoint
            });
            //if (Trl.iLabels == 0) Trl.SearchLabelsM("No Labels");
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
            //combinationsList.Clear();
            distinctUnits = Enumerable.Empty<string>();
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
            //ListUp($"APIKey = {API_Req.APIKey}");
            //ListUp($"myTrelloToken = {API_Req.myTrelloToken}");
            //ListUp($"boardCode = {API_Req.boardCode}");
            //try
            //{
            //    ReadOnlySpan<byte> s_readToEnd_stringUtf8 = Encoding.UTF8.GetBytes(API_Req.ReadToEnd_string);
            //    var reader = new Utf8JsonReader(s_readToEnd_stringUtf8);
            //    while (reader.Read())
            //    {
            //        JsonTokenType tokenType;
            //        tokenType = reader.TokenType;
            //        switch (tokenType)
            //        {
            //            case JsonTokenType.StartObject:
            //                break;
            //            case JsonTokenType.PropertyName:
            //                if (reader.ValueTextEquals(s_nameUtf8))
            //                {
            //                    try { NameTokenM(reader); }
            //                    catch (Exception e)
            //                    {
            //                        ListUp(e.Message);
            //                        return;
            //                    }
            //                }
            //                else if (reader.ValueTextEquals(s_UrlUtf8))
            //                {
            //                    try { ShortUrlTokenM(reader); }
            //                    catch (Exception e)
            //                    {
            //                        ListUp(e.Message);
            //                        return;
            //                    }
            //                }
            //                else if (reader.ValueTextEquals(s_cardRoleUtf8))
            //                {
            //                    try { CardRoleTokenM(reader); }
            //                    catch (Exception e)
            //                    {
            //                        ListUp(e.Message);
            //                        return;
            //                    }
            //                    ClearCardM();
            //                }
            //                break;
            //        }
            //    }
            //    labels.Sort();
            //}
            //catch (Exception e)
            //{
            //    ListUp(e.Message);
            //    return;
            //}
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
    }
}
