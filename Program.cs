using System.Text;

namespace ProjectClosureToolWinFormsNET6
{
    internal static class Program
    {
        ///// <summary>
        ///// Главная точка входа для приложения.
        ///// </summary>
        //[STAThread]
        public static readonly byte[] s_nameUtf8 = Encoding.UTF8.GetBytes("name");
        public static readonly byte[] s_UrlUtf8 = Encoding.UTF8.GetBytes("shortUrl");
        public static readonly byte[] s_badgesUtf8 = Encoding.UTF8.GetBytes("badges");
        public static readonly byte[] s_cardRoleUtf8 = Encoding.UTF8.GetBytes("cardRole");

        public const string c_boardURL = "https://trello.com/1/boards/";
        public const string с_boardCode = "dXURQTbH";
        public const string c_APIKey = "4b02fbde8c00369dc53e25222e864941";
        public const string c_MyTrelloToken = "717ed29e99fcd032275052b563319915f7ce0ec975c5a2abcd965ddd2cf91b07";

        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Form1());
        }
    }
}