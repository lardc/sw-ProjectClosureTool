using System.Linq;
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
using System.Windows.Markup;

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

        private void DataGridLabels_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            int iInputIgnore = DataGridLabels.Items.IndexOf(DataGridLabels.CurrentItem);
            int iColumnIndex = DataGridLabels.Items.IndexOf(DataGridLabels.CurrentItem);
            //((WindowBind)DataContext).CombinationsDownloadCheckbox(iInputIgnore, iColumnIndex);
            //((WindowBind)DataContext).CombinationsDownload();
            if ((iInputIgnore <= Download.labelsList.Count) && (iColumnIndex == 1) && iInputIgnore >= 0)
            {
                foreach (TrelloObjectLabels aLabel in Download.labelsList)
                    if (aLabel.CardID.Equals(iInputIgnore))
                    {
                        if (Download.CheckIgnored(aLabel.CardLabel))
                        {
                            Download.ignoredLabelsList.Remove(aLabel);
                            DataGridLabels.CurrentItem = false;
                            //DataGridLabels.CurrentCell.Value = false;
                            if (Download.ignoredLabelsList.Count == 0)
                                Download.ignoredLabelsListFilled = false;
                        }
                        else
                        {
                            Download.ignoredLabelsList.Add(aLabel);
                            DataGridLabels.CurrentItem = true;
                            //DataGridLabels.CurrentCell.Value = true;
                            Download.ignoredLabelsListFilled = true;
                        }
                    }
                ((WindowBind)DataContext).CombinationsDownloadCheckbox(iInputIgnore, iColumnIndex);
            }
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            ((WindowBind)DataContext).CombinationsDownload();
        }

        private void Button_Click_2(object sender, RoutedEventArgs e)
        {
            ((WindowBind)DataContext).Export();
        }

        private void DataGridResults_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {

        }

        private void Button_Click_3(object sender, RoutedEventArgs e)
        {
            ((WindowBind)DataContext).Results();
        }

        private void DataGrid_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {

        }

        private void DataGridCombinations_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {

        }
    }
}