using CashFlowApp.UI.model;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Transactions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace CashFlowApp.UI
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void ChooseExpensesFile(object sender, RoutedEventArgs e)
        {
            ExpensesFilePath.Text = ChooseFile();
        }

        private void ChooseIncomesFile(object sender, RoutedEventArgs e)
        {
            IncomesFilePath.Text = ChooseFile();
        }

        private string ChooseFile()
        {
            Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();
            if (dlg.ShowDialog() == true)
            {
                return dlg.FileName;
            }
            return string.Empty;
        }

        private void GenerateExcel(object sender, RoutedEventArgs e)
        {
            CashFlowFacade cashFlowFacade = new CashFlowFacade();
            cashFlowFacade.GenerateOutput(new CashFlowFilesModel()
            {
                ExpensesFilePath = ExpensesFilePath.Text
            });

          
        }

       
    }
}
