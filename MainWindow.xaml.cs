using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace TestMVVM
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private int i = 1;
        //private string textValue = "";
        //private int randomValue1;
        //private int randomValue2;
        //private int sum;


        public MainWindow()
        {
            InitializeComponent();
            DataContext = new WindowBind();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            ((WindowBind)DataContext).WindowBindButton(i);
            i++;
        }

        private void textBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            
        }

        private void DataGrid_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {

        }

        //public class Table
        //{
        //    public string Text { get; set; }
        //    //public int Random1 { get; set; }
        //    //public int Random2 { get; set; }
        //    //public int Sum { get; set; }
        //}

        //public void TextBox_TextChanged(object sender, TextChangedEventArgs e)
        //{
        //    textValue = $"{i}. {e}";
        //    i++;
        //    var rand = new Random();
        //    randomValue1 = rand.Next();
        //    randomValue2 = rand.Next();
        //    sum = randomValue1 + randomValue2;
        //}
    }
}