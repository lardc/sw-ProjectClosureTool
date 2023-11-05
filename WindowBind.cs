using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;

namespace TestMVVM
{
    internal class WindowBind : INotifyPropertyChanged
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

        public void WindowBindButton(int i)
        {
            Models.Clear();
            for (int j = 0; j<100; j++) 
            {
                Models.Add(new Model(Text, j));
            }
            //Models.Add(new Model { Text = $"{i}. {text}" });
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

        private string? text;

        public string Text { get => text; set => SetProperty(ref text, value); }
    }
}
