using System.Security.Policy;
using System.Text.Json;
using System.Text;
using System.Data;
using System.Windows.Forms;
using Microsoft.VisualBasic;
using System.ComponentModel;
using Microsoft.VisualBasic.Devices;
using System.Linq;
using OfficeOpenXml;

namespace ProjectClosureToolWinFormsNET6
{
    public partial class Form1 : Form
    {
        private static readonly byte[] s_nameUtf8 = Encoding.UTF8.GetBytes("name");
        private static readonly byte[] s_UrlUtf8 = Encoding.UTF8.GetBytes("shortUrl");
        private static readonly byte[] s_cardRoleUtf8 = Encoding.UTF8.GetBytes("cardRole");

        private static List<TrelloObject> cards = new List<TrelloObject>();
        static List<TrelloObjectLabels> labels = new List<TrelloObjectLabels>();
        static List<TrelloObjectLabels> labelsList = new List<TrelloObjectLabels>();
        static List<TrelloObjectLabels> ignoredLabelsList = new List<TrelloObjectLabels>();
        private static List<string> units = new List<string>();
        static IEnumerable<string> distinctUnits = new List<string>();
        private static List<string> distinctUnitsList;
        static List<string> combinationsList = new List<string>();
        static List<string> combinationsListI = new List<string>();
        static IEnumerable<string> distinctCombinations = new List<string>();
        static IEnumerable<string> distinctCombinationsI = new List<string>();
        private static List<string> distinctCombinationsList;
        private static List<string> distinctCombinationsListI = new List<string>();
        private static bool labelsListFilled;
        private static bool ignoredLabelsListFilled;
        private static bool unitsListFilled;
        private static bool distinctCombinationsListFilled = false;
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
        private static List<string> selectedUnits = new List<string>();
        private static List<string> selectedCombinations = new List<string>();

        private static bool selectedCombinationsFilled = false;
        private static bool sumsListFilled = false;
        private static bool isIgnoredDGV2 = false;
        private static bool isCheckedDGV3 = false;
        private static bool isCheckedDGV4 = false;

        static List<TrelloObjectSums> sums = new List<TrelloObjectSums>();

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
        private void DGV2Up(string mss)
        {
            dataGridView2_new_string(mss);
        }
        private void DGV3Up(string mss)
        {
            dataGridView3_new_string(mss);
        }
        private void DGV4Up(string mss)
        {
            dataGridView4_new_string(mss);
        }
        private void DGV5Up(string mssU, string mssC, string mssSE, string mssSP)
        {
            dataGridView5_new_string(mssU, mssC, mssSE, mssSP);
        }
        public static DataTable table = new DataTable();
        public static DataTable tableDGV2 = new DataTable();
        public static DataTable tableDGV3 = new DataTable();
        public static DataTable tableDGV4 = new DataTable();
        public static DataTable tableDGV5 = new DataTable();

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

        public struct MyOutDGV2
        {
            public string strOUT;
            public bool isIgn;
            public MyOutDGV2(string _strOUT, bool _isIgn)
            {
                strOUT = _strOUT;
                isIgn = _isIgn;
            }
        }
        public static List<MyOutDGV2> myOutsDGV2 = new List<MyOutDGV2>();

        public struct MyOutDGV3
        {
            public string strOUT;
            public bool isCh;
            public MyOutDGV3(string _strOUT, bool _isCh)
            {
                strOUT = _strOUT;
                isCh = _isCh;
            }
        }
        public static List<MyOutDGV3> myOutsDGV3 = new List<MyOutDGV3>();

        public struct MyOutDGV4
        {
            public string strOUT;
            public bool isCh;
            public MyOutDGV4(string _strOUT, bool _isCh)
            {
                strOUT = _strOUT;
                isCh = _isCh;
            }
        }
        public static List<MyOutDGV4> myOutsDGV4 = new List<MyOutDGV4>();
        public struct MyOutDGV5
        {
            public string unit;
            public string combination;
            public string sumEst;
            public string sumPoint;
            public MyOutDGV5(string _unit, string _combination, string _sumEst, string _sumPoint)
            {
                unit = _unit;
                combination = _combination;
                sumEst = _sumEst;
                sumPoint = _sumPoint;
            }
        }
        public static List<MyOutDGV5> myOutsDGV5 = new List<MyOutDGV5>();

        public Form1()
        {
            InitializeComponent();
        }

        public void Form1_Load(object sender, EventArgs e)
        {
            dataGridView1.AllowUserToAddRows = false;
            dataGridView5.AllowUserToAddRows = false;
            table.Columns.Add("N п/п", typeof(int));
            table.Columns.Add("Response", typeof(string));
            dataGridView1_start_string("---------------Старт-----------");
            tableDGV2.Columns.Add("Label", typeof(string));
            tableDGV2.Columns.Add("Ignored", typeof(bool));
            tableDGV3.Columns.Add("Unit", typeof(string));
            tableDGV3.Columns.Add("Selected", typeof(bool));
            tableDGV4.Columns.Add("Selected", typeof(bool));
            tableDGV4.Columns.Add("Combination", typeof(string));
            tableDGV5.Columns.Add("Unit", typeof(string));
            tableDGV5.Columns.Add("Combination", typeof(string));
            tableDGV5.Columns.Add("Sum Est.", typeof(string));
            tableDGV5.Columns.Add("Sum P.", typeof(string));
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
            ignoredLabelsList.Clear();
            ignoredLabelsListFilled = false;
            selectedUnits.Clear();
            selectedCombinations.Clear();
            dataGridView2_clear_string();
            dataGridView3_clear_string();
            dataGridView4_clear_string();
            dataGridView5_clear_string();
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
            LabelsList();
            if (unitsListFilled)
            {
                distinctUnits = units.Distinct();
                distinctUnitsList = distinctUnits.ToList();
                distinctUnitsList.Sort();
                iUnit = 1;
                foreach (string aUnit in distinctUnitsList)
                {
                    isCheckedDGV3 = selectedUnits.Contains(aUnit);
                    DGV3Up($"{iUnit}. {aUnit}");
                    isCheckedDGV3 = false;
                    iUnit++;
                }
                iUnit = 0;
            }
            LabelCombinationsI();
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

        private void dataGridView2_new_string(string mss)
        {
            int k = myOutsDGV2.Count;
            myOutsDGV2.Add(new MyOutDGV2(mss, isIgnoredDGV2));
            tableDGV2.Rows.Add(myOutsDGV2[k].strOUT, myOutsDGV2[k].isIgn);
            dataGridView2.DataSource = tableDGV2;
            //if (k > 1)
            //    dataGridView2.CurrentCell = dataGridView2[0, k - 1];
        }

        private void dataGridView3_new_string(string mss)
        {
            int k = myOutsDGV3.Count;
            myOutsDGV3.Add(new MyOutDGV3(mss, isCheckedDGV3));
            tableDGV3.Rows.Add(myOutsDGV3[k].strOUT, myOutsDGV3[k].isCh);
            dataGridView3.DataSource = tableDGV3;
        }

        private void dataGridView2_clear_string()
        {
            tableDGV2.Rows.Clear();
            dataGridView2.DataSource = tableDGV2;
        }

        private void dataGridView3_clear_string()
        {
            tableDGV3.Rows.Clear();
            dataGridView3.DataSource = tableDGV3;
        }

        private void dataGridView4_clear_string()
        {
            tableDGV4.Rows.Clear();
            dataGridView4.DataSource = tableDGV4;
        }

        private void dataGridView5_clear_string()
        {
            tableDGV5.Rows.Clear();
            dataGridView5.DataSource = tableDGV5;
        }

        private void dataGridView4_new_string(string mss)
        {
            int k = myOutsDGV4.Count;
            myOutsDGV4.Add(new MyOutDGV4(mss, isCheckedDGV4));
            tableDGV4.Rows.Add(myOutsDGV4[k].isCh, myOutsDGV4[k].strOUT);
            dataGridView4.DataSource = tableDGV4;
        }

        private void dataGridView5_new_string(string mssU, string mssC, string mssSE, string mssSP)
        {
            int k = myOutsDGV5.Count;
            myOutsDGV5.Add(new MyOutDGV5(mssU, mssC, mssSE, mssSP));
            tableDGV5.Rows.Add(myOutsDGV5[k].unit, myOutsDGV5[k].combination, myOutsDGV5[k].sumEst, myOutsDGV5[k].sumPoint);
            dataGridView5.DataSource = tableDGV5;
            //if (k > 0)
            //    dataGridView5.CurrentCell = dataGridView5[0, k - 1];
        }

        private void button2_Click(object sender, EventArgs e)
        {
            //dataGridView2.Visible = true;
            //dataGridView3.Visible = false;
            //dataGridView4.Visible = false;
            //dataGridView5.Visible = false;

            //int count = dataGridView2.Rows.Count;
            //if (count > 0)
            //    for (int i = 0; i < count; i++)
            //    {
            //        dataGridView2.Rows.Remove(dataGridView2.Rows[0]);
            //    }
            //LabelsList();
        }

        private void LabelsList()
        {
            if (!labelsListFilled)
                ListUp("Ярлыков нет. Выполните команду ввода кода доски.");
            else
            {
                foreach (TrelloObjectLabels aLabel in labelsList)
                {
                    isIgnoredDGV2 = CheckIgnored(aLabel.CardLabel);
                    DGV2Up(aLabel.ToString());
                    isIgnoredDGV2 = false;
                }
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

        // Список блоков
        private void button3_Click(object sender, EventArgs e)
        {
            //dataGridView2.Visible = false;
            //dataGridView3.Visible = true;
            //dataGridView4.Visible = false;
            //dataGridView5.Visible = false;

            //int count = dataGridView3.Rows.Count;
            //if (count > 0)
            //    for (int i = 0; i < count; i++)
            //    {
            //        dataGridView3.Rows.Remove(dataGridView3.Rows[0]);
            //    }
            //if (unitsListFilled)
            //{
            //    distinctUnits = units.Distinct();
            //    distinctUnitsList = distinctUnits.ToList();
            //    distinctUnitsList.Sort();
            //    iUnit = 1;
            //    foreach (string aUnit in distinctUnitsList)
            //    {
            //        isCheckedDGV3 = iSelectedUnit == distinctUnitsList.IndexOf(aUnit);
            //        DGV3Up($"{iUnit}. {aUnit}");
            //        isCheckedDGV3 = false;
            //        iUnit++;
            //    }
            //    iUnit = 0;
            //}
        }

        public void LabelCombinations()
        {
            dataGridView4_clear_string();
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
                    DGV4Up($"{iCombination}. {aCombination}");
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
            dataGridView4_clear_string();
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
            distinctCombinationsListFilled = true;
            try
            {
                iCombination = 1;
                foreach (string aCombination in distinctCombinationsListI)
                {
                    isCheckedDGV4 = selectedCombinations.Contains(aCombination);
                    DGV4Up($"{iCombination}. {aCombination}");
                    isCheckedDGV4 = false;
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
            //LabelCombinations();
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
                dataGridView2_clear_string();
                LabelsList();
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
            //int count = dataGridView2.Rows.Count;
            //if (count > 0)
            //    for (int i = 0; i < count - 1; i++)
            //    {
            //        dataGridView2.Rows.Remove(dataGridView2.Rows[0]);
            //    }
            //if (!labelsListFilled)
            //    ListUp("Ярлыков нет. Выполните команду ввода кода доски.");
            //else
            //{
            //    LabelsList();
            //    Application.DoEvents();
            //    AddIgnore();
            //}
        }

        // Просмотр списка игнорируемых ярлыков
        private void button6_Click(object sender, EventArgs e)
        {
            //int count = dataGridView2.Rows.Count;
            //if (count > 0)
            //    for (int i = 0; i < count - 1; i++)
            //    {
            //        dataGridView2.Rows.Remove(dataGridView2.Rows[0]);
            //    }
            //ListUp("Перечень игнорируемых ярлыков:");
            //IgnoredLabelsList();
        }

        // Очистка списка игнорируемых ярлыков
        private void button7_Click(object sender, EventArgs e)
        {
            //ignoredLabelsList.Clear();
            //ignoredLabelsListFilled = false;
            //int count = dataGridView2.Rows.Count;
            //if (count > 0)
            //    for (int i = 0; i < count; i++)
            //    {
            //        dataGridView2.Rows.Remove(dataGridView2.Rows[0]);
            //    }
            //LabelsList();
        }

        // Список комбинаций ярлыков (игнорируемые не учитываются)
        private void button8_Click(object sender, EventArgs e)
        {
            //LabelCombinationsI();
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
                    dataGridView2_clear_string();
                    LabelsList();
                    Application.DoEvents();
                }
                else
                {
                    foreach (TrelloObjectLabels aLabel in temp)
                        if (ignoredLabelsList.Contains(aLabel))
                            ignoredLabelsList.Remove(aLabel);
                    dataGridView2_clear_string();
                    LabelsList();
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
            //if (!ignoredLabelsListFilled)
            //    ListUp("Игнорируемых ярлыков нет");
            //else
            //{
            //    ListUp("Удаление ярлыков по номеру");
            //    ListUp("Для выхода введите 0");
            //    //IgnoredLabelsList();
            //    Application.DoEvents();
            //    DeleteIgnoredLabel();
            //    int count = dataGridView2.Rows.Count;
            //    if (count > 0)
            //        for (int i = 0; i < count - 1; i++)
            //        {
            //            dataGridView2.Rows.Remove(dataGridView2.Rows[0]);
            //        }
            //    LabelsList();
            //}
        }

        private void Sum(int i, int j)
        {
            sumEstimate = 0;
            sumPoint = 0;
            string unit = distinctUnitsList.ElementAt(i);
            string combination = distinctCombinationsListI.ElementAt(j - 1);
            int sumT = 0;
            foreach (TrelloObject aCard in cards)
                if (aCard.CardUnit.Equals(unit) && aCard.LabelCombinationI.Equals(combination))
                {
                    sumEstimate += aCard.CardEstimate;
                    sumPoint += aCard.CardPoint;
                    sumT++;
                    sums.Add(new TrelloObjectSums()
                    {
                        CardUnit = unit,
                        LabelCombinationI = combination,
                        SumEstimate = sumEstimate,
                        SumPoint = sumPoint
                    });
                }
            if (sumT > 0)
            {
                DGV5Up(unit, combination, sumEstimate.ToString(), sumPoint.ToString());
                sumsListFilled = true;
            }
        }

        // Суммарные оценки для выбранных блоков для всех комбинаций ярлыков
        private void button10_Click(object sender, EventArgs e)
        {
            if (!labelsListFilled || !unitsListFilled)
                ListUp("Выполните команду ввода кода доски");
            else
            {
                dataGridView5_clear_string();
                LabelCombinationsI();
                foreach (string aUnit in selectedUnits)
                    foreach (string aCombination in distinctCombinationsListI)
                        Sum(distinctUnitsList.IndexOf(aUnit), distinctCombinationsListI.IndexOf(aCombination) + 1);
            }
        }

        // Суммарные оценки для выбранных блоков для выбранных комбинаций ярлыков
        private void button11_Click(object sender, EventArgs e)
        {
            if (!labelsListFilled || !unitsListFilled)
                ListUp("Выполните команду ввода кода доски");
            else
            {
                dataGridView5_clear_string();
                foreach (string aUnit in selectedUnits)
                    foreach (string aCombination in selectedCombinations)
                        Sum(distinctUnitsList.IndexOf(aUnit), distinctCombinationsListI.IndexOf(aCombination) + 1);

                //dataGridView3.Visible = false;
                //dataGridView5.Visible = true;
            }
        }

        private void dataGridView2_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            int iInputIgnore = dataGridView2.CurrentCell.RowIndex;
            if (iInputIgnore <= labelsList.Count)
            {
                foreach (TrelloObjectLabels aLabel in labelsList)
                    if (aLabel.CardID.Equals(iInputIgnore))
                    {
                        if (CheckIgnored(aLabel.CardLabel))
                        {
                            ignoredLabelsList.Remove(aLabel);
                            if (ignoredLabelsList.Count == 0)
                                ignoredLabelsListFilled = false;
                        }
                        else
                        {
                            ignoredLabelsList.Add(aLabel);
                            ignoredLabelsListFilled = true;
                        }
                    }
                selectedCombinations.Clear();
                dataGridView4_clear_string();
            }
            LabelCombinationsI();
            dataGridView5_clear_string();
            foreach (string aUnit in selectedUnits)
            {
                if (selectedCombinationsFilled)
                    foreach (string aCombination in selectedCombinations)
                        Sum(distinctUnitsList.IndexOf(aUnit), distinctCombinationsListI.IndexOf(aCombination) + 1);
                else
                    foreach (string aCombination in distinctCombinationsListI)
                        Sum(distinctUnitsList.IndexOf(aUnit), distinctCombinationsListI.IndexOf(aCombination) + 1);
            }
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        // Сохранение списка игнорируемых ярлыков
        private void button12_Click(object sender, EventArgs e)
        {
            //int count = dataGridView2.Rows.Count;
            //if (count > 0)
            //    for (int i = 0; i < count; i++)
            //    {
            //        dataGridView2.Rows.Remove(dataGridView2.Rows[0]);
            //    }
            //LabelsList();
        }

        private void dataGridView3_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            int iInputUnit = dataGridView3.CurrentCell.RowIndex;
            if (iInputUnit < distinctUnitsList.Count)
            {
                dataGridView5_clear_string();
                if (selectedUnits.Contains(distinctUnitsList.ElementAt(iInputUnit)))
                    selectedUnits.Remove(distinctUnitsList.ElementAt(iInputUnit));
                else selectedUnits.Add(distinctUnitsList.ElementAt(iInputUnit));
                LabelCombinationsI();
                foreach (string aUnit in selectedUnits)
                {
                    if (selectedCombinationsFilled)
                        foreach (string aCombination in selectedCombinations)
                            Sum(distinctUnitsList.IndexOf(aUnit), distinctCombinationsListI.IndexOf(aCombination) + 1);
                    else
                        foreach (string aCombination in distinctCombinationsListI)
                            Sum(distinctUnitsList.IndexOf(aUnit), distinctCombinationsListI.IndexOf(aCombination) + 1);
                }
            }
        }

        private void dataGridView4_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            int iInputCombination = dataGridView4.CurrentCell.RowIndex;
            if (iInputCombination < distinctCombinationsListI.Count)
            {
                if (selectedCombinations.Contains(distinctCombinationsListI.ElementAt(iInputCombination)))
                {
                    selectedCombinations.Remove(distinctCombinationsListI.ElementAt(iInputCombination));
                    if (selectedCombinations.Count == 0)
                        selectedCombinationsFilled = false;
                }
                else
                {
                    selectedCombinations.Add(distinctCombinationsListI.ElementAt(iInputCombination));
                    selectedCombinationsFilled = true;
                }
                dataGridView5_clear_string();
                foreach (string aUnit in selectedUnits)
                {
                    if (selectedCombinationsFilled)
                        foreach (string aCombination in selectedCombinations)
                            Sum(distinctUnitsList.IndexOf(aUnit), distinctCombinationsListI.IndexOf(aCombination) + 1);
                    else
                        foreach (string aCombination in distinctCombinationsListI)
                            Sum(distinctUnitsList.IndexOf(aUnit), distinctCombinationsListI.IndexOf(aCombination) + 1);
                }
            }
        }

        private void dataGridView5_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void FillExcel()
        {
            DateTime DTnow = DateTime.Now;
            string DTyear = DTnow.Year.ToString();
            string DTmonth = DTnow.Month.ToString();
            string DTday = DTnow.Day.ToString();
            string DThour = DTnow.Hour.ToString();
            string DTminute = DTnow.Minute.ToString();
            string DTsecond = DTnow.Second.ToString();
            string fileName = $"{API_Req.boardCode}_D{DTyear}-{DTmonth}-{DTday}_T{DThour}-{DTminute}-{DTsecond}.xlsx";
            using (ExcelPackage excel_result = new ExcelPackage())
            {
                excel_result.Workbook.Worksheets.Add("Estimations");
                excel_result.Workbook.Worksheets.Add("Points");
                FileInfo fout = new FileInfo(fileName);
                if (fout.Exists)
                {
                    fout.Delete();
                    fout = new FileInfo(@"{API_Req.boardCode}_D{DTyear}-{DTmonth}-{DTday}_T{DThour}-{DTminute}-{DTsecond}.xlsx");
                }
                excel_result.SaveAs(fout);
                var estimateWorksheet = excel_result.Workbook.Worksheets["Estimations"];
                var pointWorksheet = excel_result.Workbook.Worksheets["Points"];
                for (int iU = 1; iU <= distinctUnitsList.Count; iU++)
                {
                    estimateWorksheet.Cells[iU + 1, 1].Style.WrapText = true;
                    pointWorksheet.Cells[iU + 1, 1].Style.WrapText = true;
                    estimateWorksheet.Cells[iU + 1, 1].Value = distinctUnitsList[iU - 1];
                    pointWorksheet.Cells[iU + 1, 1].Value = distinctUnitsList[iU - 1];
                    for (int iC = 1; iC <= distinctCombinationsListI.Count; iC++)
                    {
                        estimateWorksheet.Cells[1, iC + 1].Style.WrapText = true;
                        pointWorksheet.Cells[1, iC + 1].Style.WrapText = true;
                        if (iU == distinctUnitsList.Count)
                        {
                            estimateWorksheet.Cells[1, iC + 1].Value = distinctCombinationsListI[iC - 1];
                            pointWorksheet.Cells[1, iC + 1].Value = distinctCombinationsListI[iC - 1];
                        }
                        estimateWorksheet.Column(iC).Width = 19;
                        estimateWorksheet.Column(iC + 1).Width = 19;
                        pointWorksheet.Column(iC).Width = 19;
                        pointWorksheet.Column(iC + 1).Width = 19;
                    }
                }
                excel_result.SaveAs(fout);
                for (int iU = 1; iU <= distinctUnitsList.Count; iU++)
                {
                    for (int iC = 1; iC <= distinctCombinationsListI.Count; iC++)
                    {
                        string unit = distinctUnitsList.ElementAt(iU - 1);
                        string combination = distinctCombinationsListI.ElementAt(iC - 1);
                        estimateWorksheet.Cells[iU + 1, iC + 1].Style.WrapText = true;
                        pointWorksheet.Cells[iU + 1, iC + 1].Style.WrapText = true;
                        foreach (TrelloObjectSums aSum in sums)
                            if (aSum.CardUnit.Equals(unit) && aSum.LabelCombinationI.Equals(combination))
                            {
                                estimateWorksheet.Cells[iU + 1, iC + 1].Value = aSum.SumEstimate;
                                pointWorksheet.Cells[iU + 1, iC + 1].Value = aSum.SumPoint;
                            }
                    }
                }
                excel_result.SaveAs(fout);
            }
        }

        private void button13_Click(object sender, EventArgs e)
        {
            if (labelsListFilled && unitsListFilled && distinctCombinationsListFilled)
            {
                foreach (string aUnit in distinctUnitsList)
                    foreach (string aCombination in distinctCombinationsListI)
                        Sum(distinctUnitsList.IndexOf(aUnit), distinctCombinationsListI.IndexOf(aCombination) + 1);
                FillExcel();
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