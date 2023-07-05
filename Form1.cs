using System.Security.Policy;
using System.Text.Json;
using System.Text;
using System.Data;
using System.Windows.Forms;
using Microsoft.VisualBasic;

namespace ProjectClosureToolWinFormsNET6
{
    public partial class Form1 : Form
    {
        private static readonly byte[] s_nameUtf8 = Encoding.UTF8.GetBytes("name");
        private static readonly byte[] s_UrlUtf8 = Encoding.UTF8.GetBytes("shortUrl");
        private static readonly byte[] s_cardRoleUtf8 = Encoding.UTF8.GetBytes("cardRole");

        public static List<TrelloObject> cards = new List<TrelloObject>();
        static List<TrelloObjectLabels> labels = new List<TrelloObjectLabels>();
        static List<TrelloObjectLabels> labelsList = new List<TrelloObjectLabels>();
        static List<TrelloObjectLabels> ignoredLabelsList = new List<TrelloObjectLabels>();
        public static List<string> units = new List<string>();
        static IEnumerable<string> distinctUnits = new List<string>();
        public static List<string> distinctUnitsList;
        static List<string> combinationsList = new List<string>();
        static List<string> combinationsListI = new List<string>();
        static IEnumerable<string> distinctCombinations = new List<string>();
        static IEnumerable<string> distinctCombinationsI = new List<string>();
        private static List<string> distinctCombinationsList;
        private static List<string> distinctCombinationsListI;
        private static bool labelsListFilled;
        private static bool ignoredLabelsListFilled;
        private static bool unitsListFilled;
        private static int iLabels;
        private static string[] cardLabels = new string[1000];

        public static string currentCardURL;
        public static string currentCardUnit;
        public static string currentCardName;
        public static double currentCardEstimate;
        public static double currentCardPoint;
        public static double sumEstimate;
        public static double sumPoint;
        public static int iInputIgnore;
        private static int iDeleteIgnore;
        private static int iUnit;
        private static int iCombination;

        private void EMessage(string eM)
        {
            int s_err = eM.IndexOf("error", 0);
            int s_dot = eM.IndexOf(".", 0);
            string s_mess = eM.Substring(s_err, s_dot - s_err).Trim();
            ListUp(s_mess);
        }

        private void ListUp(string mss)
        {
            dataGridView1_new_string(mss);
        }
        public static DataTable table = new DataTable();

        //-------------------------------------
        public struct MyOut
        {
            public int numStr;
            public string strOUT;
            public MyOut(int _numStr, string _strOUT)
            {
                numStr = _numStr;
                strOUT = _strOUT;
            }
        }
        public static List<MyOut> myOuts = new List<MyOut>();

        public Form1()
        {
            InitializeComponent();
        }

        public void Form1_Load(object sender, EventArgs e)
        {
            dataGridView1.AllowUserToAddRows = false;
            table.Columns.Add("N п/п", typeof(int));
            table.Columns.Add("Response", typeof(string));
            dataGridView1_start_string("---------------Старт-----------");
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            API_Req.boardURL = textBox1.Text;
            Application.DoEvents();
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            API_Req.boardCode = textBox2.Text;
            Application.DoEvents();
        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {
            if (textBox3.Text == Program.c_APIKey)
            {
                API_Req.APIKey = textBox3.Text;
                textBox3.Text = "Default";
            }
            if (textBox3.Text != "Default")
            { API_Req.APIKey = textBox3.Text; }
        }

        private void textBox4_TextChanged(object sender, EventArgs e)
        {
            if (textBox4.Text == Program.c_MyTrelloToken)
            {
                API_Req.myTrelloToken = textBox4.Text;
                textBox4.Text = "Default";
            }
            if (textBox4.Text != "Default")
            { API_Req.myTrelloToken = textBox4.Text; }
        }

        public void ClearLabels()
        {
            labelsList.Clear();
            iLabels = 0;
        }

        public void ClearCardM()
        {
            currentCardURL = "";
            currentCardUnit = "";
            currentCardName = "";
            currentCardEstimate = 0;
            currentCardPoint = 0;
            for (int i = 0; i < 20; i++)
                cardLabels[i] = "";
            iLabels = 0;
        }

        // Получение кода для работы с конкретной доской
        private void ReadBoard()
        {
            string CardFilter = "/cards/open";
            string CardFields = "&fields=id,badges,dateLastActivity,idBoard,idLabels,idList,idShort,labels,limits,name,shortLink,shortUrl,cardRole,url&limit=1000";

            try
            {
                API_Req.RequestAsync(API_Req.APIKey, API_Req.myTrelloToken, CardFilter, CardFields, API_Req.boardCode);
            }
            catch (Exception e)
            {
                if (e.Message.Contains("400"))
                {
                    EMessage(e.Message);
                    ListUp("Недопустимый url доски");
                    return;
                }
                if (e.Message.Contains("401"))
                {
                    EMessage(e.Message);
                    ListUp("Нет доступа к доске \nkt - ввод ключа и токена\n");
                    return;
                }
                if (e.Message.Contains("404"))
                {
                    EMessage(e.Message);
                    ListUp("Доска не найдена");
                    return;
                }
            }
        }

        private void BoardM()
        {
            ClearCardM();
            cards.Clear();
            units.Clear();
            labels.Clear();
            combinationsList.Clear();
            distinctUnits = Enumerable.Empty<string>();
            ListUp($"APIKey = {API_Req.APIKey}");
            ListUp($"myTrelloToken = {API_Req.myTrelloToken}");
            ListUp($"boardCode = {API_Req.boardCode}");
            try
            {
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
                            if (reader.ValueTextEquals(s_nameUtf8))
                            {
                                try { NameTokenM(reader); }
                                catch (Exception e)
                                {
                                    ListUp(e.Message);
                                    return;
                                }
                            }
                            else if (reader.ValueTextEquals(s_UrlUtf8))
                            {
                                try { ShortUrlTokenM(reader); }
                                catch (Exception e)
                                {
                                    ListUp(e.Message);
                                    return;
                                }
                            }
                            else if (reader.ValueTextEquals(s_cardRoleUtf8))
                            {
                                try { CardRoleTokenM(reader); }
                                catch (Exception e)
                                {
                                    ListUp(e.Message);
                                    return;
                                }
                                ClearCardM();
                            }
                            break;
                    }
                }
                labels.Sort();
            }
            catch (Exception e)
            {
                ListUp(e.Message);
                return;
            }
        }

        private void SearchUnitValuesM(string rr)
        {
            if (rr.Length > 0)
            {
                int iDot = rr.IndexOf('.', 0); // позиция точки в строке rr
                int iOpeningParenthesis = rr.IndexOf('(', 0); // позиция открывающейся круглой скобки в строке rr
                int iClosingParenthesis = rr.IndexOf(')', 0); // позиция закрывающейся круглой скобки в строке rr
                int iOpeningBracket = rr.IndexOf('[', 0); // позиция открывающейся квадратной скобки в строке rr
                int iClosingBracket = rr.IndexOf(']', 0); // позиция закрывающейся квадратной скобки в строке rr

                if (iDot >= 0)
                {
                    if (iOpeningParenthesis > iDot && iClosingParenthesis > iDot)
                    {
                        string s_uiro = rr.Substring(iOpeningParenthesis + 1, iClosingParenthesis - iOpeningParenthesis - 1).Trim();
                        double d_ior = double.Parse(s_uiro);
                        try { currentCardEstimate = d_ior; }
                        catch (Exception e)
                        {
                            ListUp(e.Message);
                            return;
                        }
                    }
                    if (iOpeningBracket > iDot && iClosingBracket > iDot)
                    {
                        string s_uisq = rr.Substring(iOpeningBracket + 1, iClosingBracket - iOpeningBracket - 1).Trim();
                        double d_isq = double.Parse(s_uisq);
                        try { currentCardPoint = d_isq; }
                        catch (Exception e)
                        {
                            ListUp(e.Message);
                            return;
                        }
                    }
                    string s_unit = rr.Substring(0, iDot).Trim();
                    try
                    {
                        currentCardUnit = s_unit;
                        units.Add(currentCardUnit);
                        unitsListFilled = true;
                    }
                    catch (Exception e)
                    {
                        ListUp(e.Message);
                        return;
                    }
                    string s_name = "";
                    if (iOpeningParenthesis >= 0 && iOpeningBracket >= 0)
                    {
                        if (rr.Substring(iDot + 1, Math.Min(iOpeningParenthesis, iOpeningBracket) - iDot - 1).Trim().Length > 0)
                            s_name = rr.Substring(iDot + 1, Math.Min(iOpeningParenthesis, iOpeningBracket) - iDot - 1).Trim();
                    }
                    else if (iOpeningParenthesis < 0 && iOpeningBracket < 0)
                    {
                        s_name = rr.Substring(iDot + 1, rr.Length - iDot - 1).Trim();
                    }
                    else if (rr.Substring(iDot + 1, Math.Max(iOpeningParenthesis, iOpeningBracket) - iDot - 1).Trim().Length > 0)
                        s_name = rr.Substring(iDot + 1, Math.Max(iOpeningParenthesis, iOpeningBracket) - iDot - 1).Trim();
                    try { currentCardName = s_name; }
                    catch (Exception e)
                    {
                        ListUp(e.Message);
                        return;
                    }
                }
                else currentCardName = rr;
            }
        }

        public void SearchLabelsM(string rr)
        {
            cardLabels[iLabels] = rr;
            iLabels++;
        }

        private void NameTokenM(Utf8JsonReader reader)
        {
            if (reader.GetString().StartsWith("name"))
            {
                reader.Read();
                if (reader.CurrentDepth.Equals(2))
                    try { SearchUnitValuesM(reader.GetString().ToString()); }
                    catch (Exception e)
                    {
                        ListUp(e.Message);
                        return;
                    }
                else if (reader.CurrentDepth.Equals(4))
                    try { SearchLabelsM(reader.GetString().ToString()); }
                    catch (Exception e)
                    {
                        ListUp(e.Message);
                        return;
                    }
            }
        }

        public void ShortUrlTokenM(Utf8JsonReader reader)
        {
            if (reader.GetString().StartsWith("shortUrl"))
            {
                reader.Read();
                currentCardURL = reader.GetString().ToString();
            }
        }

        public void CardRoleTokenM(Utf8JsonReader reader)
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
            for (int i = 0; i < iLabels; i++)
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

        // Обработка
        private void button1_Click(object sender, EventArgs e)
        {
            ClearLabels();
            labelsListFilled = false;
            unitsListFilled = false;
            try
            { API_Req.boardCode = textBox2.Text; }
            catch (Exception ex)
            {
                ListUp(ex.Message);
                return;
            }
            ReadBoard();
            BoardM();
            ListUp("Обработка завершена");
        }

        private void dataGridView1_start_string(string mss)
        {
            myOuts.Add(new MyOut(0, mss));
            table.Rows.Clear();
            table.Rows.Add(0, myOuts[0].strOUT);
            dataGridView1.DataSource = table;
        }


        private void dataGridView1_new_string(string mss)
        {
            int k = myOuts.Count;
            myOuts.Add(new MyOut(k, mss));
            table.Rows.Add(myOuts[k].numStr, myOuts[k].strOUT);
            dataGridView1.DataSource = table;
            dataGridView1.CurrentCell = dataGridView1[0, k];
        }

        private void button2_Click(object sender, EventArgs e)
        {
            ListUp("Список ярлыков:");
            LabelsList();
        }

        private void LabelsList()
        {
            if (!labelsListFilled)
                ListUp("Ярлыков нет. Выполните команду ввода кода доски.");
            else
            {
                foreach (TrelloObjectLabels aLabel in labelsList)
                    if (!CheckIgnored(aLabel.CardLabel))
                        ListUp(aLabel.ToString());
            }
        }

        private void UnitsList()
        {
            if (!unitsListFilled)
                ListUp("Блоков нет. Выполните команду ввода кода доски.");
            else
            {
                distinctUnits = units.Distinct();
                distinctUnitsList = distinctUnits.ToList();
                distinctUnitsList.Sort();
                iUnit = 1;
                foreach (string aUnit in distinctUnitsList)
                {
                    ListUp($"{iUnit}. {aUnit}");
                    iUnit++;
                }
                iUnit = 0;
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            ListUp("Список блоков:");
            //distinctUnits = units.Distinct();
            //distinctUnitsList = distinctUnits.ToList();
            //distinctUnitsList.Sort();
            UnitsList();
        }

        public void LabelCombinations()
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
            try
            {
                iCombination = 1;
                foreach (string aCombination in distinctCombinationsList)
                {
                    ListUp($"{iCombination}. {aCombination}");
                    iCombination++;
                }
                iCombination = 0;
            }
            catch (Exception e)
            {
                ListUp(e.Message);
                return;
            }
        }

        private void LabelCombinationsI()
        {
            combinationsListI.Clear();
            distinctCombinationsI = Enumerable.Empty<string>();
            labels.Sort();
            for (int i = 0; i < cards.Count; i++)
            {
                string sCombination = "";
                foreach (TrelloObjectLabels aLabel in labels)
                    if (aLabel.CardID.Equals(i) && !ignoredLabelsList.Contains(aLabel))
                        sCombination += $"{aLabel.CardLabel}        ";
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
            try
            {
                iCombination = 1;
                foreach (string aCombination in distinctCombinationsListI)
                {
                    ListUp($"{iCombination}. {aCombination}");
                    iCombination++;
                }
                iCombination = 0;
            }
            catch (Exception e)
            {
                ListUp(e.Message);
                return;
            }
        }

        // Список комбинаций ярлыков
        private void button4_Click(object sender, EventArgs e)
        {
            //combinationsList.Clear();
            //distinctCombinations = Enumerable.Empty<string>();
            //labels.Sort();
            //for (int i = 0; i < cards.Count; i++)
            //{
            //    string sCombination = "";
            //    foreach (TrelloObjectLabels aLabel in labels)
            //        if (aLabel.CardID.Equals(i))
            //            sCombination += $"{aLabel.CardLabel}        ";
            //    foreach (TrelloObject aCard in cards)
            //        if (aCard.CardID.Equals(i) && sCombination != "")
            //        {
            //            aCard.LabelCombination = sCombination;
            //            combinationsList.Add(sCombination);
            //        }
            //}
            //distinctCombinations = combinationsList.Distinct();
            //distinctCombinationsList = distinctCombinations.ToList();
            //distinctCombinationsList.Sort();
            ListUp("Список комбинаций ярлыков:");
            LabelCombinations();
        }

        private bool CheckLabels(string rr)
        {
            bool isLabel = false;
            foreach (TrelloObjectLabels aLabel in labelsList)
            {
                if (aLabel.CardLabel.Equals(rr))
                    isLabel = true;
            }
            return isLabel;
        }

        private bool CheckIgnored(string rr)
        {
            bool isIgnored = false;
            foreach (TrelloObjectLabels aLabel in ignoredLabelsList)
            {
                if (aLabel.CardLabel.Equals(rr))
                    isIgnored = true;
            }
            return isIgnored;
        }

        private void AddIgnore()
        {
            try { iInputIgnore = int.Parse(Interaction.InputBox("Введите номер ярлыка")); }
            catch (Exception e)
            {
                ListUp(e.Message);
                return;
            }
            while (iInputIgnore != 0)
            {
                foreach (TrelloObjectLabels aLabel in labelsList)
                    if (aLabel.CardID.Equals(iInputIgnore - 1))
                    {
                        if (!CheckLabels(aLabel.CardLabel) || iInputIgnore > labelsList.Count)
                            ListUp($"Нет ярлыка с номером <{iInputIgnore}>. Введите другой номер.");
                        else
                        {
                            if (CheckIgnored(aLabel.CardLabel))
                            {
                                foreach (TrelloObjectLabels label in labelsList)
                                    if (label.CardID.Equals(iInputIgnore - 1))
                                        ListUp($"Ярлык <{label.CardLabel}> уже содержится в списке игнорируемых");
                            }
                            else
                            {
                                ignoredLabelsList.Add(aLabel);
                                ignoredLabelsListFilled = true;
                            }
                        }
                    }
                ListUp("Перечень игнорируемых ярлыков:");
                IgnoredLabelsList();
                Application.DoEvents();
                try { iInputIgnore = int.Parse(Interaction.InputBox("Введите номер ярлыка")); }
                catch (Exception e)
                {
                    ListUp(e.Message);
                    return;
                }
            }
            foreach (TrelloObject aCard in cards)
                aCard.LabelCombinationI = "";
        }

        // Добавление ярлыков к списку игнорируемых
        private void button5_Click(object sender, EventArgs e)
        {
            if (!labelsListFilled)
                ListUp("Ярлыков нет. Выполните команду ввода кода доски.");
            else
            {
                LabelsList();
                Application.DoEvents();
                AddIgnore();
            }
        }

        private void IgnoredLabelsList()
        {
            if (!ignoredLabelsListFilled)
                ListUp("Игнорируемых ярлыков нет");
            else
                foreach (TrelloObjectLabels aLabel in ignoredLabelsList)
                {
                    if (!CheckLabels(aLabel.CardLabel))
                        ListUp($"Игнорируемый ярлык <{aLabel.CardLabel}> отсутствует в списке ярлыков. Перезагрузите доску.");
                    else ListUp(aLabel.ToString());
                }
        }

        // Просмотр списка игнорируемых ярлыков
        private void button6_Click(object sender, EventArgs e)
        {
            ListUp("Перечень игнорируемых ярлыков:");
            IgnoredLabelsList();
        }

        // Очистка списка игнорируемых ярлыков
        private void button7_Click(object sender, EventArgs e)
        {
            ignoredLabelsList.Clear();
            ignoredLabelsListFilled = false;
            IgnoredLabelsList();
        }

        // Список комбинаций ярлыков (игнорируемые не учитываются)
        private void button8_Click(object sender, EventArgs e)
        {
            //combinationsListI.Clear();
            //distinctCombinationsI = Enumerable.Empty<string>();
            //labels.Sort();
            //for (int i = 0; i < cards.Count; i++)
            //{
            //    string sCombination = "";
            //    foreach (TrelloObjectLabels aLabel in labels)
            //        if (aLabel.CardID.Equals(i) && !ignoredLabelsList.Contains(aLabel))
            //            sCombination += $"{aLabel.CardLabel}        ";
            //    foreach (TrelloObject aCard in cards)
            //        if (aCard.CardID.Equals(i) && sCombination != "")
            //        {
            //            aCard.LabelCombination = sCombination;
            //            combinationsListI.Add(sCombination);
            //        }
            //}
            //distinctCombinationsI = combinationsListI.Distinct();
            //distinctCombinationsListI = distinctCombinationsI.ToList();
            //distinctCombinationsListI.Sort();
            ListUp("Список комбинаций ярлыков (игнорируемые не учитываются):");
            LabelCombinationsI();
        }

        private void DeleteIgnoredLabel()
        {
            try { iDeleteIgnore = int.Parse(Interaction.InputBox("Введите номер удаляемого из списка игнорируемых ярлыка")); }
            catch (Exception ex)
            {
                ListUp(ex.Message);
                return;
            }
            Application.DoEvents();
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
                    ListUp($"Нет игнорируемого ярлыка с номером <{iDeleteIgnore}>. Введите другой номер.");
                if (ignoredLabelsList.Count < 1)
                {
                    ignoredLabelsListFilled = false;
                    iDeleteIgnore = 0;
                    ListUp("Перечень игнорируемых ярлыков:");
                    IgnoredLabelsList();
                    Application.DoEvents();
                }
                else
                {
                    foreach (TrelloObjectLabels aLabel in temp)
                        if (ignoredLabelsList.Contains(aLabel))
                            ignoredLabelsList.Remove(aLabel);
                    ListUp("Перечень игнорируемых ярлыков:");
                    IgnoredLabelsList();
                    Application.DoEvents();
                    try { iDeleteIgnore = int.Parse(Interaction.InputBox("Введите номер удаляемого из списка игнорируемых ярлыка")); }
                    catch (Exception ex)
                    {
                        ListUp(ex.Message);
                        return;
                    }
                }
            }
        }

        // Удаление ярлыков из списка игнорируемых
        private void button9_Click(object sender, EventArgs e)
        {
            if (!ignoredLabelsListFilled)
                ListUp("Игнорируемых ярлыков нет");
            else
            {
                ListUp("Удаление ярлыков по номеру");
                ListUp("Для выхода введите 0");
                IgnoredLabelsList();
                Application.DoEvents();
                DeleteIgnoredLabel();
            }
        }

        private void Sum(int i, int j)
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
                }
            if (sumT > 0)
            {
                ListUp($"Блок: {unit}");
                ListUp($"Комбинация ярлыков:");
                ListUp($"{combination}");
                ListUp($"Суммарное оценочное значение: ({sumEstimate})");
                ListUp($"Суммарное реальное значение: [{sumPoint}].");
            }
        }

        // Суммарные оценки для конкретнго блока для всех комбинаций ярлыков
        private void button10_Click(object sender, EventArgs e)
        {
            if (!labelsListFilled || !unitsListFilled)
                ListUp("Выполните команду ввода кода доски");
            else
            {
                ListUp("Список комбинаций ярлыков (игнорируемые не учитываются):");
                LabelCombinationsI();
                Application.DoEvents();
                ListUp("Список блоков:");
                UnitsList();
                Application.DoEvents();
                try { iUnit = int.Parse(Interaction.InputBox("Введите номер блока")); }
                catch (Exception ex)
                {
                    ListUp(ex.Message);
                    return;
                }
                if (iUnit > distinctUnitsList.Count || iUnit < 1)
                    ListUp("Введите корректный номер блока");
                else for (int i = 0; i < distinctCombinationsI.ToList().Count; i++)
                        Sum(iUnit, i + 1);
            }
        }

        // Суммарные оценки для конкретнго блока для конкретной комбинации ярлыков
        private void button11_Click(object sender, EventArgs e)
        {
            if (!labelsListFilled || !unitsListFilled)
                ListUp("Выполните команду ввода кода доски");
            else
            {
                ListUp("Список блоков:");
                UnitsList();
                Application.DoEvents();
                try { iUnit = int.Parse(Interaction.InputBox("Введите номер блока")); }
                catch (Exception ex)
                {
                    ListUp(ex.Message);
                    return;
                }
                if (iUnit > distinctUnitsList.Count || iUnit < 1)
                    ListUp("Введите корректный номер блока");
                else
                {
                    ListUp("Список комбинаций ярлыков (игнорируемые не учитываются):");
                    LabelCombinationsI();
                    Application.DoEvents();
                    try { iCombination = int.Parse(Interaction.InputBox("Введите номер комбинации")); }
                    catch (Exception ex)
                    {
                        ListUp(ex.Message);
                        return;
                    }
                    if (iCombination > distinctCombinationsListI.Count || iCombination < 1)
                        { ListUp("Введите корректный номер комбинации"); }
                    else Sum(iUnit, iCombination);
                }
            }
        }
    }
    public class API_Req
    {
        private static string readToEnd_string;
        public static string boardURL = Program.c_boardURL;
        public static string boardCode = Program.с_boardCode;
        public static string APIKey = Program.c_APIKey;
        public static string myTrelloToken = Program.c_MyTrelloToken;
        public static string CardFilter = "/cards/open";
        public static string CardFields = "&fields=id,badges,dateLastActivity,idBoard,idLabels,idList,idShort,labels,limits,name,shortLink,shortUrl,cardRole,url&limit=1000";

        public static string ReadToEnd_string { get => readToEnd_string; set => readToEnd_string = value; }

        static readonly HttpClient client = new HttpClient();

        /// Ключ и токен для авторизации
        /// Запрос карточек доски
        /// Получение ответа в виде потока
        /// Запись ответа в память виде строки
        /// Кодировка в UTF-8
        public static async Task RequestAsync(string APIKey, string MyTrelloToken, string CardFilter, string CardFields, string boardCode)
        {
            System.Net.WebRequest reqGET = System.Net.WebRequest.Create(boardURL + boardCode + CardFilter + "/?key=" + APIKey + "&token=" + MyTrelloToken + CardFields);
            System.Net.WebResponse resp = reqGET.GetResponse();
            System.IO.Stream stream = resp.GetResponseStream();
            System.IO.StreamReader read_stream = new(stream);
            ReadToEnd_string = read_stream.ReadToEnd();
        }
    }
}