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

namespace Pauls_Contest_Mailer
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow() => InitializeComponent();

        private Random rnd = new Random();

        private SpreadsheetAnalysis analysis = null;

        private void BtnBrowse_Click(object sender, RoutedEventArgs e)
        {
            Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog
            {
                DefaultExt = ".txt",
                Filter = "Excel-Dateien (*.xls, *.xlsx)|*.xls;*.xlsx"
            };

            bool? result = dlg.ShowDialog();
            if (result == true)
            {
                // Open document
                string filename = dlg.FileName;

                analysis = new SpreadsheetAnalysis();
                analysis.ProcessWorkbook(filename);

                if (!analysis.mailNotFound)
                {
                    btnSendMails.IsEnabled = true;
                }
            }
        }

        private void BtnSendMails_Click(object sender, RoutedEventArgs e)
        {
            int total = analysis.CountMails();

            MessageBoxResult result = MessageBox.Show("Das kann dauern ;)" + Environment.NewLine + "Es werden " + total + " Emails im Intervall von " + slValue.Value + " Minuten versandt." + Environment.NewLine + "Erwartetes Ende: " + DateTime.Now.AddMinutes(total*slValue.Value) + Environment.NewLine + "Jetzt starten?", "Erwartete Dauer", MessageBoxButton.OKCancel);

            if (result == MessageBoxResult.OK)
            {
                btnSendMails.IsEnabled = false;
                btnBrowse.IsEnabled = false;

                analysis.SendMails(mailGrid, slValue.Value);
            }
        }
    }
}
