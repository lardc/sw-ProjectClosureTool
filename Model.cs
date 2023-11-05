using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO.Packaging;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TestMVVM
{
    internal class Model : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler? PropertyChanged;

        public void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        private string? text;
        public string? Text { get => text; set { text = value; OnPropertyChanged("Text"); } }

        private Random rand = new Random();

        private int randomValue1;
        public int RandomValue1 { get => randomValue1; set { randomValue1 = value; OnPropertyChanged("RandomValue1"); } }

        private int randomValue2;
        public int RandomValue2 { get => randomValue2; set { randomValue2 = value; OnPropertyChanged("RandomValue2"); } }

        private int sum;
        public int Sum { get => sum; set { sum = randomValue1 + randomValue2; OnPropertyChanged("Sum"); } }

        private int counter;
        //public int Counter { get => counter; set { counter = value; OnPropertyChanged("Counter"); } }

        public Model(string arg, int counter)
        {
            counter++;
            Text = arg + " " + counter.ToString();
            RandomValue1 = rand.Next(101);
            RandomValue2 = rand.Next(101);
            Sum = RandomValue1 + RandomValue2;
        }
    }
}
