using DocumentFormat.OpenXml.Office2010.Word;
using DocumentFormat.OpenXml.Presentation;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Database;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Math;
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
using System.Windows.Data;
using System.Windows.Input;
using static System.Net.Mime.MediaTypeNames;

namespace ProjectClosureToolMVVM
{
    internal class WindowBind: INotifyPropertyChanged
    {
        public ICommand MyCommand { get; set; }

        public event PropertyChangedEventHandler? PropertyChanged;
        public ObservableCollection<Model> labelModels = new ObservableCollection<Model>();
        public ObservableCollection<Model> unitModels = new ObservableCollection<Model>();
        public ObservableCollection<Model> combinationModels = new ObservableCollection<Model>();
        public ObservableCollection<ResultsModel> resultsModels = new ObservableCollection<ResultsModel>();

        public void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        public ObservableCollection<Model> LabelModels { get { return labelModels; } set { labelModels = value; OnPropertyChanged("LabelModels"); } }
        public ObservableCollection<Model> UnitModels { get { return unitModels; } set { unitModels = value; OnPropertyChanged("UnitModels"); } }
        public ObservableCollection<Model> CombinationModels { get { return combinationModels; } set { combinationModels = value; OnPropertyChanged("UnitModels"); } }
        public ObservableCollection<ResultsModel> ResultsModels { get { return resultsModels; } set { resultsModels = value; OnPropertyChanged("ResultModels"); } }

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
            LabelModels.Clear();
            UnitModels.Clear();
            CombinationModels.Clear();
            ResultsModels.Clear();
            if (BoardCode != API_Req.boardCode && BoardCode != "" && BoardCode != null)
                { API_Req.boardCode = BoardCode; }
            if (APIKey != "Default" && APIKey != "" && APIKey != null)
                { API_Req.APIKey = APIKey; }
            if (myTrelloToken != "Default" && myTrelloToken != "" && myTrelloToken != null)
                { API_Req.myTrelloToken = myTrelloToken; }
            API_Req.RequestAsync(API_Req.APIKey, API_Req.myTrelloToken, API_Req.CardFilter, API_Req.CardFields, API_Req.boardCode);
            Download.BoardM();
            foreach (TrelloObjectLabels aLabel in Download.labelsList)
                LabelModels.Add(new Model(aLabel.CardLabel, Download.CheckIgnored(aLabel.CardLabel)));
            if (Download.unitsListFilled)
            {
                Download.distinctUnits = Download.units.Distinct();
                Download.distinctUnitsList = Download.distinctUnits.ToList();
                Download.distinctUnitsList.Sort();
                foreach (string aUnit in Download.distinctUnitsList)
                    UnitModels.Add(new Model(aUnit, Download.selectedUnits.Contains(aUnit)));
            }
        }

        public void CombinationsDownload()
        {
            Download.ignoredLabelsList.Clear();
            Download.ignoredLabelsListFilled = false;
            CombinationModels.Clear();
            ResultsModels.Clear();
            foreach (Model aLabel in LabelModels)
                if (aLabel.IsIgnored)
                {
                    Download.ignoredLabelsList.Add(new TrelloObjectLabels() 
                    {
                        CardID=Download.ignoredLabelsList.Count,
                        CardLabel= aLabel.Label
                    });
                    Download.ignoredLabelsListFilled = true;
                }
            Download.LabelCombinationsI();
            foreach (string aCombination in Download.distinctCombinationsListI)
                CombinationModels.Add(new Model(aCombination, Download.selectedCombinations.Contains(aCombination)));
            Trl.sums.Clear();
            for (int i = 0; i < Download.distinctUnitsList.Count(); i++)
                for (int j = 0; j < Download.distinctCombinationsListI.Count(); j++)
                    Trl.Sum(i, j + 1);
        }

        public void Results()
        {
            ResultsModels.Clear();
            Download.selectedUnits.Clear();
            foreach (Model aUnit in UnitModels)
                if (aUnit.IsIgnored && !Download.selectedUnits.Contains(aUnit.Label))
                    Download.selectedUnits.Add(aUnit.Label);
            Download.selectedCombinations.Clear();
            foreach (Model aCombination in CombinationModels)
                if (aCombination.IsIgnored && !Download.selectedCombinations.Contains(aCombination.Label))
                    Download.selectedCombinations.Add(aCombination.Label);
            Trl.sums.Clear();
            for (int i = 0; i < Download.distinctUnitsList.Count(); i++)
                for (int j = 0; j < Download.distinctCombinationsListI.Count(); j++)
                    Trl.Sum(i, j + 1);
            foreach (TrelloObjectSums aSum in Trl.sums)
                if (Download.selectedUnits.Contains(aSum.CardUnit) && Download.selectedCombinations.Contains(aSum.LabelCombinationI))
                    ResultsModels.Add(new ResultsModel(aSum.CardUnit, aSum.LabelCombinationI, aSum.SumEstimate, aSum.SumPoint));
        }

        public void Export()
        {
            if (Download.labelsListFilled && Download.unitsListFilled && Download.distinctCombinationsListFilled)
            {
                ResultsModels.Clear();
                Download.selectedUnits.Clear();
                foreach (Model aUnit in UnitModels)
                    if (aUnit.IsIgnored && !Download.selectedUnits.Contains(aUnit.Label))
                        Download.selectedUnits.Add(aUnit.Label);
                Download.selectedCombinations.Clear();
                foreach (Model aCombination in CombinationModels)
                    if (aCombination.IsIgnored && !Download.selectedCombinations.Contains(aCombination.Label))
                        Download.selectedCombinations.Add(aCombination.Label);
                Trl.sums.Clear();
                foreach (string aUnit in Download.distinctUnitsList)
                    foreach (string aCombination in Download.distinctCombinationsListI)
                            Trl.Sum(Download.distinctUnitsList.IndexOf(aUnit), Download.distinctCombinationsListI.IndexOf(aCombination) + 1);
                TableResp.FillExcel();
            }
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
