using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ProjectClosureToolWinFormsNET6
{
    public partial class TrelloObject
    {
        public int CardID { get; set; }
        public string CardURL { get; set; }
        public string CardUnit { get; set; }
        public string CardName { get; set; }
        public double CardEstimate { get; set; }
        public double CardPoint { get; set; }
        public string LabelCombination { get; set; }
        public string LabelCombinationI { get; set; }
        public override string ToString()
        {
            return "CardID: " + CardID + "; CardURL: " + CardURL + "; CardUnit: " + CardUnit + "; CardName: " + CardName + "; CardEstimate: " + CardEstimate + "; CardPoint: " + CardPoint + "; LabelCombinationI: " + LabelCombinationI;
        }
    }
    public partial class TrelloObjectLabels : IComparable<TrelloObjectLabels>
    {
        public int CardID { get; set; }
        public string CardLabel { get; set; }
        public override string ToString()
        {
            return $"{CardID + 1}. {CardLabel}";
        }
        public override bool Equals(object obj)
        {
            if (obj == null) return false;
            TrelloObjectLabels objAsLabel = obj as TrelloObjectLabels;
            if (objAsLabel == null) return false;
            else return Equals(objAsLabel);
        }
        public bool Equals(TrelloObjectLabels other)
        {
            if (other == null) return false;
            return this.CardLabel.Equals(other.CardLabel);
        }
        public int SortByNameAscending(string name1, string name2)
        {
            return name1.CompareTo(name2);
        }
        public int CompareTo(TrelloObjectLabels compareLabel)
        {
            if (compareLabel == null)
                return 1;
            else
                return this.CardLabel.CompareTo(compareLabel.CardLabel);
        }
        public override int GetHashCode()
        {
            return CardID;
        }
    }
    public partial class TrelloObjectSums
    {
        public string CardUnit { get; set; }
        public string LabelCombinationI { get; set; }
        public double SumEstimate { get; set; }
        public double SumPoint { get; set; }
    }
}