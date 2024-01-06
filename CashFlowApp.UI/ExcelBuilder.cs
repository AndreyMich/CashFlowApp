using CashFlowApp.UI.model;
using SpreadsheetLight;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Linq.Expressions;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace CashFlowApp.UI
{
    internal class ExcelBuilder
    {
        private int row = 2;
        public void CreateExcel(ExcelBuilderModel model)
        {
            using (var sl = new SLDocument())
            {
                // Group expenses
                // Group expenses and add to the workbook
                AddExpensesToExcelWithSL(sl, model.Expenses, "Expenses");

                // Group incomes and add to the workbook
                //AddExpensesToExcelWithSL(sl, model.Incomes, "Incomes");

                // Save the workbook to a file
                var saveFileDialog = new Microsoft.Win32.SaveFileDialog
                {
                    Filter = "Excel files (*.xlsx)|*.xlsx",
                    FileName = "Output.xlsx"
                };
                if (saveFileDialog.ShowDialog() == true)
                {
                    sl.SaveAs(saveFileDialog.FileName);
                    MessageBox.Show("Excel file generated successfully!");
                }
            }
        }

        private void AddExpensesToExcelWithSL(SLDocument sl, List<Transaction> transactions, string sheetPrefix)
        {
            var groupedTransactions = transactions.GroupBy(a => new Tuple<int,int>(a.Date.Month,a.Date.Year));

            foreach (var transaction in groupedTransactions)//iterate transactions by month and year
            {
                string sheetName = $"{transaction.Key.Item1} {transaction.Key.Item2}";
                var transactionByAccountNameGroup = transaction.ToList().GroupBy(a => a.Name); //groups of transaction by account name
                foreach(var accountGroup in transactionByAccountNameGroup)
                {
                    List<GroupedTransaction> accountTransactions = accountGroup.Select(a => new GroupedTransaction()
                    {
                        TotalAmount = a.Amount,
                        Date = a.Date.ToShortDateString(),
                        Name = a.Name
                    }).ToList();

                    AddToSheetWithSL(sl, sheetName, accountTransactions);
                }
            }
            //var now = DateTime.Now;
            //var currentMonthTransactions = transactions.Where(t => t.Date <= now).ToList();
            //if (currentMonthTransactions.Any())
            //{
            //    var groupedCurrentMonth = currentMonthTransactions.GroupBy(t => t.Name).Select(g => new GroupedTransaction()
            //    {
            //        Name = g.Key,
            //        TotalAmount = g.Sum(t => t.Amount),
            //        Date = now.ToString("MMMM yyyy")
            //    }).ToList();

            //    AddToSheetWithSL(sl, $"{sheetPrefix} - {now:MMMM yyyy}", groupedCurrentMonth);
            //}

            //var futureTransactions = transactions.Where(t => t.Date > now).ToList();
            //var groupedFutureTransactions = futureTransactions.GroupBy(t => new { t.Date.Year, t.Date.Month, t.Name }).Select(g => new GroupedTransaction
            //{
            //    Name = g.Key.Name,
            //    TotalAmount = g.Sum(t => t.Amount),
            //    Date = $"{new DateTime(g.Key.Year, g.Key.Month, 1):MMMM yyyy}"
            //}).ToList();

            //foreach (var group in groupedFutureTransactions.GroupBy(g => g.Date))
            //{
            //    AddToSheetWithSL(sl, $"{sheetPrefix} - {group.Key}", group.ToList());
            //}


        }

        private void AddToSheetWithSL(SLDocument sl, string sheetName, List<GroupedTransaction> groupedData)
        {
            if (!IsSheetExist(sl,sheetName))
            {
                sl.AddWorksheet(sheetName);
                sl.SetCellValue(1, 2, "Name");
                sl.SetCellValue(1, 3, "Total Amount");
                sl.SetCellValue(1, 1, "Date");
            }
            
            sl.SelectWorksheet(sheetName);

           

            
            foreach (var data in groupedData)
            {
                sl.SetCellValue(row, 1, data.Name);
                sl.SetCellValue(row, 2, data.TotalAmount);
                sl.SetCellValue(row, 3, data.Date);
                row++;
            }
        }

        private bool IsSheetExist(SLDocument sl, string sheetName)
        {
            List<string> currentSheetNames = sl.GetSheetNames();
            return currentSheetNames.Contains(sheetName);
        }

        public List<Transaction> ParseExpenses(string path)
        {
            string redundantLineName = @"סה""כ";

            List<Transaction> transactions = new List<Transaction>();
            try
            {
                using (SLDocument sl = new SLDocument(path))
                {
                    int iRow = 2;  // assuming first row has headers

                    DateTime effectiveDate = DateTime.MinValue;
                    while (!string.IsNullOrEmpty(sl.GetCellValueAsString(iRow, 3)))
                    {

                        string dateCellVal = sl.GetCellValueAsString(iRow, 2);

                        if (string.IsNullOrEmpty(dateCellVal) || 
                            !string.IsNullOrEmpty(dateCellVal) && DateTime.TryParse(dateCellVal, out effectiveDate))
                        {
                            string nameCellVal = sl.GetCellValueAsString(iRow, 3);
                            string amountCellVal = sl.GetCellValueAsString(iRow, 7);
                            if (nameCellVal != redundantLineName)
                            {
                                transactions.Add(new Transaction()
                                {
                                    Name = nameCellVal,
                                    Amount = decimal.Parse(amountCellVal),
                                    Date = effectiveDate
                                });
                            }
                        }
                        
                      

                        iRow++;
                    }
                }
            }
            catch(Exception e)
            {

            }
         
            
            return transactions;
        }
    }
}
