using System.Security.Policy;
using System.Text.Json;
using System.Text;
using System.Data;

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
        private static bool labelsListFilled;
        private static int iLabels;
        private static int cardsCount;
        private static int iNone;
        private static string[] cardLabels = new string[1000];

        public static string currentCardURL;
        public static string currentCardUnit;
        public static string currentCardName;
        public static double currentCardEstimate;
        public static double currentCardPoint;

        public static void EMessage(string eM)
        {
            int s_err = eM.IndexOf("error", 0);
            int s_dot = eM.IndexOf(".", 0);
            string s_mess = eM.Substring(s_err, s_dot - s_err).Trim();
            myOuts.Add(new MyOut(0, s_mess));
        }

        private /*static*/ void ListUp(string mss)
        {
            //listBox1.BeginUpdate();
            //listBox1.Items.Add(mss);
            //listBox1.EndUpdate();
            //Application.DoEvents();
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

        //public struct MyOutN
        //{
        //    public int numStr;
        //    public string strOUT;


        //    public MyOutN(int _numStr, string _strOUT)
        //    {
        //        numStr = _numStr;
        //        strOUT = _strOUT;
        //    }
        //}
        //public List<MyOutN> myOutsN = new List<MyOutN>();

        //-------------------------------------




        public Form1()
        {
            InitializeComponent();
        }

        public void Form1_Load(object sender, EventArgs e)
        //private void Form1_Load(object sender, EventArgs e)
        {
            dataGridView1.AllowUserToAddRows = false;
            myOuts.Add(new MyOut(0, "---------------Старт-----------"));

            //DataTable table = new DataTable();
            table.Columns.Add("N п/п", typeof(int));
            table.Columns.Add("Response", typeof(string));

            for (int i = 0; i < myOuts.Count; i++)
            {
                table.Rows.Add(myOuts[i].numStr, myOuts[i].strOUT);
            }
            dataGridView1.DataSource = table;
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

        // Очистка списка игнорируемых ярлыков
        //public static void ClearIgnoredLabels()
        //{
        //    ignoredLabelsList.Clear();
        //    ignoredLabelsListFilled = false;
        //    LabelCombinationsI();
        //}

        public static void ClearLabels()
        {
            labelsList.Clear();
            iLabels = 0;
        }

        public static void ClearCardM()
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
        public static void ReadBoard()
        {
            string CardFilter = "/cards/open";
            string CardFields = "&fields=id,badges,dateLastActivity,idBoard,idLabels,idList,idShort,labels,limits,name,shortLink,shortUrl,cardRole,url&limit=1000";

            //myOuts.Add(new MyOut(0, "Start"));

            try
            {
                API_Req.RequestAsync(API_Req.APIKey, API_Req.myTrelloToken, CardFilter, CardFields, API_Req.boardCode);
            }
            catch (Exception e)
            {
                if (e.Message.Contains("400"))
                {
                    EMessage(e.Message);
                    myOuts.Add(new MyOut(0, "Недопустимый url доски"));
                    return;
                }
                if (e.Message.Contains("401"))
                {
                    EMessage(e.Message);
                    Console.Write("Нет доступа к доске \nkt - ввод ключа и токена\n");
                    return;
                }
                if (e.Message.Contains("404"))
                {
                    EMessage(e.Message);
                    myOuts.Add(new MyOut(0, "Доска не найдена"));
                    return;
                }
            }
        }

        public static void BoardM()
        {
            ClearCardM();
            cards.Clear();
            //units.Clear();
            labels.Clear();
            //combinationsList.Clear();
            //distinctUnits = Enumerable.Empty<string>();
            myOuts.Add(new MyOut(0, $"APIKey = {API_Req.APIKey}"));
            myOuts.Add(new MyOut(0, $"myTrelloToken = {API_Req.myTrelloToken}"));
            myOuts.Add(new MyOut(0, $"boardCode = {API_Req.boardCode}"));
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
                                    myOuts.Add(new MyOut(0, e.Message));
                                    return;
                                }
                            }
                            else if (reader.ValueTextEquals(s_UrlUtf8))
                            {
                                try { ShortUrlTokenM(reader); }
                                catch (Exception e)
                                {
                                    myOuts.Add(new MyOut(0, e.Message));
                                    return;
                                }
                            }
                            else if (reader.ValueTextEquals(s_cardRoleUtf8))
                            {
                                try { CardRoleTokenM(reader); }
                                catch (Exception e)
                                {
                                    myOuts.Add(new MyOut(0, e.Message));
                                    return;
                                }
                                ClearCardM();
                            }
                            break;
                    }
                }
                labels.Sort();
                //LabelCombinationsI();
            }
            catch (Exception e)
            {
                myOuts.Add(new MyOut(0, e.Message));
                return;
            }
        }

        public static void SearchUnitValuesM(string rr)
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
                            myOuts.Add(new MyOut(0, e.Message));
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
                            myOuts.Add(new MyOut(0, e.Message));
                            return;
                        }
                    }
                    string s_unit = rr.Substring(0, iDot).Trim();
                    try
                    {
                        currentCardUnit = s_unit;
                    }
                    catch (Exception e)
                    {
                        myOuts.Add(new MyOut(0, e.Message));
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
                        myOuts.Add(new MyOut(0, e.Message));
                        return;
                    }
                }
                else currentCardName = rr;
                if (currentCardEstimate == 0 && currentCardPoint == 0 && ((iDot <= Math.Min(iOpeningParenthesis, iOpeningBracket) && Math.Min(iOpeningParenthesis, iOpeningBracket) >= 0) || (Math.Max(iOpeningParenthesis, iOpeningBracket) < 0))) iNone++;
            }
        }

        public static void SearchLabelsM(string rr)
        {
            cardLabels[iLabels] = rr;
            iLabels++;
        }

        public static void NameTokenM(Utf8JsonReader reader)
        {
            if (reader.GetString().StartsWith("name"))
            {
                reader.Read();
                if (reader.CurrentDepth.Equals(2))
                    try { SearchUnitValuesM(reader.GetString().ToString()); }
                    catch (Exception e)
                    {
                        myOuts.Add(new MyOut(0, e.Message));
                        return;
                    }
                else if (reader.CurrentDepth.Equals(4))
                    try { SearchLabelsM(reader.GetString().ToString()); }
                    catch (Exception e)
                    {
                        myOuts.Add(new MyOut(0, e.Message));
                        return;
                    }
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
            //ClearIgnoredLabels();
            ClearLabels();
            labelsListFilled = false;
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

        private void LabelsList()
        //public static void LabelsList()
        {
            if (!labelsListFilled)
            {
                ListUp("Ярлыков нет. Выполните команду ввода кода доски.");
                string message = "Ярлыков нет. Выполните команду ввода кода доски.";
                string caption = labelsListFilled.ToString();
                MessageBoxButtons buttons = MessageBoxButtons.YesNo;
                DialogResult result;
                result = MessageBox.Show(message, caption, buttons);
                if (result == System.Windows.Forms.DialogResult.Yes)
                    this.Close();
            }
            else
            {
                foreach (TrelloObjectLabels aLabel in labelsList)
                    //if (!CheckIgnored(aLabel.CardLabel))
                    ListUp(aLabel.ToString());
            }
        }
        private void dataGridView1_new_string(string mss)
        {
            int k = myOuts.Count;
            myOuts.Add(new MyOut(k, mss));
            //myOuts.Add(new MyOut(2, "Старт 2"));

            //DataTable table = new DataTable();
            //table.Columns.Add("N п/п", typeof(int));
            //table.Columns.Add("Response", typeof(string));

            table.Rows.Clear();
            for (int i = 0; i < myOuts.Count; i++)
            {
                table.Rows.Add(myOuts[i].numStr, myOuts[i].strOUT);
            }
            dataGridView1.DataSource = table;

        }

        private void button2_Click(object sender, EventArgs e)
        {
            ListUp("Список ярлыков:");
            LabelsList();
        }
    }
    public static class API_Req
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