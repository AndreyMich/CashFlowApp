using SpreadsheetLight;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace CashFlowApp.UI
{
    internal class ExcelBuilder
    {
        public void CreateExcel(List<Transaction> expenses, List<Transaction> incomes)
        {
            using (var sl = new SLDocument())
            {
                // Group expenses
                // Group expenses and add to the workbook
                GroupAndAddToExcelWithSL(sl, expenses, "Expenses");

                // Group incomes and add to the workbook
                GroupAndAddToExcelWithSL(sl, incomes, "Incomes");

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

        private void GroupAndAddToExcelWithSL(SLDocument sl, List<Transaction> transactions, string sheetPrefix)
        {
            var now = DateTime.Now;
            var currentMonthTransactions = transactions.Where(t => t.Date <= now).ToList();
            if (currentMonthTransactions.Any())
            {
                var groupedCurrentMonth = currentMonthTransactions.GroupBy(t => t.Name).Select(g => new GroupedTransaction()
                {
                    Name = g.Key,
                    TotalAmount = g.Sum(t => t.Amount),
                    Date = now.ToString("MMMM yyyy")
                }).ToList();

                AddToSheetWithSL(sl, $"{sheetPrefix} - {now:MMMM yyyy}", groupedCurrentMonth);
            }

            var futureTransactions = transactions.Where(t => t.Date > now).ToList();
            var groupedFutureTransactions = futureTransactions.GroupBy(t => new { t.Date.Year, t.Date.Month, t.Name }).Select(g => new GroupedTransaction
            {
                Name = g.Key.Name,
                TotalAmount = g.Sum(t => t.Amount),
                Date = $"{new DateTime(g.Key.Year, g.Key.Month, 1):MMMM yyyy}"
            }).ToList();

            foreach (var group in groupedFutureTransactions.GroupBy(g => g.Date))
            {
                AddToSheetWithSL(sl, $"{sheetPrefix} - {group.Key}", group.ToList());
            }
        }

        private void AddToSheetWithSL(SLDocument sl, string sheetName, List<GroupedTransaction> groupedData)
        {
            if (!sl.AddWorksheet(sheetName))
            {
                // Handle case where sheet name might already exist
                // This is just a simple example, you might want to add a suffix or handle in another way
                sheetName += "_1";
                sl.AddWorksheet(sheetName);
            }

            sl.SelectWorksheet(sheetName);

            sl.SetCellValue(1, 1, "Name");
            sl.SetCellValue(1, 2, "Total Amount");
            sl.SetCellValue(1, 3, "Date");

            int row = 2;
            foreach (var data in groupedData)
            {
                sl.SetCellValue(row, 1, data.Name);
                sl.SetCellValue(row, 2, data.TotalAmount);
                sl.SetCellValue(row, 3, data.Date);
                row++;
            }
        }

        public List<Transaction> ParseFile(string path)
        {
            List<Transaction> transactions = new List<Transaction>();

            using (SLDocument sl = new SLDocument(path))
            {
                int iRow = 2;  // assuming first row has headers

                while (!string.IsNullOrEmpty(sl.GetCellValueAsString(iRow, 1)))
                {
                    transactions.Add(new Transaction
                    {
                        Amount = decimal.Parse(sl.GetCellValueAsString(iRow, 1)),   // assuming column 1 has Amount
                        Date =sl.GetCellValueAsDateTime(iRow, 2),   // assuming column 2 has Date
                        Name = sl.GetCellValueAsString(iRow, 3)                    // assuming column 3 has Name
                    });

                    iRow++;
                }
            }

            return transactions;
        }
    }
}
