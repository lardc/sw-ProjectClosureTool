using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ProjectClosureToolV2
{
    public partial class API_Req
    {
        public static string boardURL;
        public static string boardCode;
        public static string APIKey;
        public static string myTrelloToken;

        public static string readToEnd_string;

        /// Ключ и токен для авторизации
        /// Запрос карточек доски
        /// Получение ответа в виде потока
        /// Запись ответа в память виде строки
        /// Кодировка в UTF-8
        public static void Request(string APIKey, string myTrelloToken, string CardFilter, string CardFields, string boardCode)
        {
            System.Net.WebRequest reqGET = System.Net.WebRequest.Create(boardURL + boardCode + CardFilter + "/?key=" + APIKey + "&token=" + myTrelloToken + CardFields);
            System.Net.WebResponse resp = reqGET.GetResponse();
            System.IO.Stream stream = resp.GetResponseStream();
            System.IO.StreamReader read_stream = new(stream);
            readToEnd_string = read_stream.ReadToEnd();
        }
    }
}