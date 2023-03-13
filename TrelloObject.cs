using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ProjectClosureToolV2
{
    public partial class TrelloObject
    {
        //// ввод доски в память
        //public static string[] cardURL = new string[iAllUnits];
        //public static string[] cardUnit = new string[iAllUnits];
        //public static int iCard = -1;
        //public static double[,] cardsValues = new double[iAllUnits, 2];
        //public static string[,] cardsLabels = new string[iAllUnits, 10];

        //// ввод текущей карточки
        //public static string currentCardURL;
        //public static string currentCardUnit;
        //public static string currentCardName;
        ////public static double[] cardValues = new double[2];
        ////public static string[] cardLabels = new string[10];

        public int CardID { get; set; }
        public string CardURL { get; set; }
        public string CardUnit { get; set; }
        public string CardName { get; set; }
        public double CardEstimate { get; set; }
        public double CardPoint { get; set; }
        //public string CardsLabels { get; set; }

        public override string ToString()
        {
            return "CardID: " + CardID + "; CardURL: " + CardURL + "; CardUnit: " + CardUnit + "; CardName: " + CardName + "; CardEstimate: " + CardEstimate + "; CardPoint: " + CardPoint;
        }
    }
    public partial class TrelloObjectLabels
    {
        public int CardID { get; set; }
        public string CardLabel { get; set; }
        public override string ToString()
        {
            return "CardID: " + CardID + "; Label: " + CardLabel;
        }
    }
}
