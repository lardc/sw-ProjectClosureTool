using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;
using System.ComponentModel;
using System.Windows.Input;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Math;

namespace ProjectClosureToolMVVM
{
    public class API_Req
    {
        public static readonly byte[] s_nameUtf8 = Encoding.UTF8.GetBytes("name");
        public static readonly byte[] s_UrlUtf8 = Encoding.UTF8.GetBytes("shortUrl");
        public static readonly byte[] s_badgesUtf8 = Encoding.UTF8.GetBytes("badges");
        public static readonly byte[] s_cardRoleUtf8 = Encoding.UTF8.GetBytes("cardRole");

        private static string readToEnd_string;
        public static string boardURL = "https://trello.com/1/boards/";
        public static string boardCode = "Be1vZgJd";
        //public static string boardCode = "dXURQTbH";
        public static string APIKey = "4b02fbde8c00369dc53e25222e864941";
        public static string myTrelloToken = "ATTA86486fcb69688e946cd7697952aedd037533786170a1840c5081a6e631b5878aCF905ED7";
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
            string fileName = "response.json";
            File.WriteAllText(fileName, ReadToEnd_string);
        }
    }
    internal class Model: INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler? PropertyChanged;
        public void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
        private string label;
        public string Label { get => label; set { label = value; OnPropertyChanged("Label"); } }
        private bool ignored;

        public bool IsIgnored {
            get { return ignored; }
            set {
                if (ignored == value) return;
                ignored = value;
                OnPropertyChanged("IsIgnored");
            }
        }

        public Model(string arg, bool ch)
        {
            Label = arg;
            IsIgnored = ch;
        }
    }

    internal class ResultsModel : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler? PropertyChanged;
        public void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
        private string unit;
        public string Unit { get => unit; set { unit = value; OnPropertyChanged("Unit"); } }

        private string combination;
        public string Combination { get => combination; set { combination = value; OnPropertyChanged("Combination"); } }

        private double sumEst;
        public double SumEst { get => sumEst; set { sumEst = value; OnPropertyChanged("SumEst"); } }

        private double sumP;
        public double SumP { get => sumP; set { sumP = value; OnPropertyChanged("SumP"); } }

        public ResultsModel(string arg1, string arg2, double argE, double argP)
        {
            Unit = arg1;
            Combination = arg2;
            SumEst = argE;
            SumP = argP;
        }
    }
}
