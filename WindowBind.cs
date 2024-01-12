using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Configuration;
using System.Linq;
using System.Net.Http;
using System.Reflection.Emit;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;
using static System.Net.Mime.MediaTypeNames;

namespace ProjectClosureToolMVVM
{
    internal class WindowBind: INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler? PropertyChanged;
        ObservableCollection<Model> labelModels = new ObservableCollection<Model>();
        ObservableCollection<Model> unitModels = new ObservableCollection<Model>();
        ObservableCollection<Model> combinationModels = new ObservableCollection<Model>();

        public void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        public ObservableCollection<Model> LabelModels { get { return labelModels; } set { labelModels = value; OnPropertyChanged("LabelModels"); } }
        public ObservableCollection<Model> UnitModels { get { return unitModels; } set { unitModels = value; OnPropertyChanged("UnitModels"); } }
        public ObservableCollection<Model> CombinationModels { get { return combinationModels; } set { combinationModels = value; OnPropertyChanged("UnitModels"); } }

        public WindowBind()
        {
            LabelModels = new ObservableCollection<Model>();
            UnitModels = new ObservableCollection<Model>();
            CombinationModels = new ObservableCollection<Model>();
        }

        private string? boardCode;
        public string BoardCode { get => boardCode; set { boardCode = value; OnPropertyChanged("BoardCode"); } }
        private string? apiKey;
        public string APIKey { get => apiKey; set { apiKey = value; OnPropertyChanged("APIKey"); } }
        private string? myTrelloToken;
        public string MyTrelloToken { get => myTrelloToken; set { myTrelloToken = value; OnPropertyChanged("MyTrelloToken"); } }

        public void WindowBindDownload()
        {
            Download.ClearLabels();
            Download.labelsListFilled = false;
            Download.unitsListFilled = false;
            Download.ignoredLabelsList.Clear();
            Download.ignoredLabelsListFilled = false;
            Download.selectedUnits.Clear();
            Download.selectedCombinations.Clear();
            Download.selectedCombinationsFilled = false;
            LabelModels.Clear();
            if (BoardCode != API_Req.boardCode && BoardCode != "" && BoardCode != null)
                { API_Req.boardCode = BoardCode; }
            if (APIKey != "Default" && APIKey != "" && APIKey != null)
                { API_Req.APIKey = APIKey; }
            if (myTrelloToken != "Default" && myTrelloToken != "" && myTrelloToken != null)
                { API_Req.myTrelloToken = myTrelloToken; }
            API_Req.RequestAsync(API_Req.APIKey, API_Req.myTrelloToken, API_Req.CardFilter, API_Req.CardFields, API_Req.boardCode);
            Download.BoardM();
            if (Download.unitsListFilled)
            {
                Download.distinctUnits = Download.units.Distinct();
                Download.distinctUnitsList = Download.distinctUnits.ToList();
                Download.distinctUnitsList.Sort();
                foreach (string aUnit in Download.distinctUnitsList)
                    UnitModels.Add(new Model(aUnit, Download.selectedUnits.Contains(aUnit)));
            }
            foreach (TrelloObjectLabels aLabel in Download.labelsList)
                LabelModels.Add(new Model(aLabel.CardLabel, Download.CheckIgnored(aLabel.CardLabel)));
            Download.LabelCombinationsI();
            foreach (string aCombination in Download.distinctCombinationsListI)
                CombinationModels.Add(new Model(aCombination, Download.selectedCombinations.Contains(aCombination)));
        }

        protected bool SetProperty<T>(ref T field, T newValue, [CallerMemberName] string? propertyName = null)
        {
            if (!Equals(field, newValue))
            {
                field = newValue;
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
                return true;
            }
            return false;
        }
    }
}
