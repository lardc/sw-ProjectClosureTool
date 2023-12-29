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
        ObservableCollection<Model> models = new ObservableCollection<Model>();

        public void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        public ObservableCollection<Model> Models { get { return models; } set { models = value; OnPropertyChanged("Models"); } }

        public WindowBind()
        {
            Models = new ObservableCollection<Model>();
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
            //unitsListFilled = false;
            Download.ignoredLabelsList.Clear();
            Download.ignoredLabelsListFilled = false;
            //selectedUnits.Clear();
            //selectedCombinations.Clear();
            Models.Clear();
            if (BoardCode != API_Req.boardCode && BoardCode != "" && BoardCode != null)
                { API_Req.boardCode = BoardCode; }
            if (APIKey != "Default" && APIKey != "" && APIKey != null)
                { API_Req.APIKey = APIKey; }
            if (myTrelloToken != "Default" && myTrelloToken != "" && myTrelloToken != null)
                { API_Req.myTrelloToken = myTrelloToken; }
            API_Req.RequestAsync(API_Req.APIKey, API_Req.myTrelloToken, API_Req.CardFilter, API_Req.CardFields, API_Req.boardCode);
            Download.BoardM();
            foreach (TrelloObjectLabels aLabel in Download.labelsList)
                Models.Add(new Model(aLabel.CardLabel, Download.CheckIgnored(aLabel.CardLabel)));
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
