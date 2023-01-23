using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TrlConsCs
{
    public partial class API_Req
    {
        public static string boardURL;
        public static string boardCode;
        public static string APIKey;
        public static string MyTrelloToken;

        public static string ReadToEnd_string;

        /// Ключ и токен для авторизации
        /// Запрос карточек доски
        /// Получение ответа в виде потока
        /// Запись ответа в память виде строки
        /// Кодировка в UTF-8
        public static void Request(string APIKey, string MyTrelloToken, string CardFilter, string CardFields, string boardCode)
        {
            System.Net.WebRequest reqGET = System.Net.WebRequest.Create(boardURL + boardCode + CardFilter + "/?key=" + APIKey + "&token=" + MyTrelloToken + CardFields);
            System.Net.WebResponse resp = reqGET.GetResponse();
            System.IO.Stream stream = resp.GetResponseStream();
            System.IO.StreamReader read_stream = new(stream);
            ReadToEnd_string = read_stream.ReadToEnd();
        }
    }
}
