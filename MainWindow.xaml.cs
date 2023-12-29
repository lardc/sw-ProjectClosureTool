using System.Net.Http;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace ProjectClosureToolMVVM
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            DataContext = new WindowBind();
        }

        // Board code
        private void TextBox_TextChanged_1(object sender, TextChangedEventArgs e)
        {

        }

        // Key
        private void TextBox_TextChanged_2(object sender, TextChangedEventArgs e)
        {

        }

        // Token
        private void TextBox_TextChanged_3(object sender, TextChangedEventArgs e)
        {

        }

        // Download
        private void Button_Click(object sender, RoutedEventArgs e)
        {
            ((WindowBind)DataContext).WindowBindDownload();
        }
    }
}