using CashFlowApp.UI.model;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml;
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
            string filePath = GetSaveFilePath();

            using (SpreadsheetDocument document = SpreadsheetDocument.Create(filePath, SpreadsheetDocumentType.Workbook))
            {
                WorkbookPart workbookPart = document.AddWorkbookPart();
                workbookPart.Workbook = new Workbook();

                Sheets sheets = document.WorkbookPart.Workbook.AppendChild(new Sheets());

                // Group expenses and add to the workbook
                AddExpensesToExcelWithOpenXML(workbookPart, sheets, model.Expenses, "Expenses");

                // Optionally, do the same for incomes
                // AddExpensesToExcelWithOpenXML(workbookPart, sheets, model.Incomes, "Incomes");

                workbookPart.Workbook.Save();
            }

            MessageBox.Show("Excel file generated successfully!");
        }

        //private void AddExpensesToExcelWithSL(SLDocument sl, List<Transaction> transactions, string sheetPrefix)
        //{
        //    var groupedTransactions = transactions.GroupBy(a => new Tuple<int,int>(a.Date.Month,a.Date.Year));

        //    foreach (var transaction in groupedTransactions)//iterate transactions by month and year
        //    {
        //        string sheetName = $"{transaction.Key.Item1} {transaction.Key.Item2}";
        //        var transactionByAccountNameGroup = transaction.ToList().GroupBy(a => a.Name); //groups of transaction by account name
        //        foreach(var accountGroup in transactionByAccountNameGroup)
        //        {
        //            List<GroupedTransaction> accountTransactions = accountGroup.Select(a => new GroupedTransaction()
        //            {
        //                TotalAmount = a.Amount,
        //                Date = a.Date.ToShortDateString(),
        //                Name = a.Name
        //            }).ToList();

        //            AddToSheetWithSL(sl, sheetName, accountTransactions);
        //        }
        //    }
        //    //var now = DateTime.Now;
        //    //var currentMonthTransactions = transactions.Where(t => t.Date <= now).ToList();
        //    //if (currentMonthTransactions.Any())
        //    //{
        //    //    var groupedCurrentMonth = currentMonthTransactions.GroupBy(t => t.Name).Select(g => new GroupedTransaction()
        //    //    {
        //    //        Name = g.Key,
        //    //        TotalAmount = g.Sum(t => t.Amount),
        //    //        Date = now.ToString("MMMM yyyy")
        //    //    }).ToList();

        //    //    AddToSheetWithSL(sl, $"{sheetPrefix} - {now:MMMM yyyy}", groupedCurrentMonth);
        //    //}

        //    //var futureTransactions = transactions.Where(t => t.Date > now).ToList();
        //    //var groupedFutureTransactions = futureTransactions.GroupBy(t => new { t.Date.Year, t.Date.Month, t.Name }).Select(g => new GroupedTransaction
        //    //{
        //    //    Name = g.Key.Name,
        //    //    TotalAmount = g.Sum(t => t.Amount),
        //    //    Date = $"{new DateTime(g.Key.Year, g.Key.Month, 1):MMMM yyyy}"
        //    //}).ToList();

        //    //foreach (var group in groupedFutureTransactions.GroupBy(g => g.Date))
        //    //{
        //    //    AddToSheetWithSL(sl, $"{sheetPrefix} - {group.Key}", group.ToList());
        //    //}


        //}

        private void AddExpensesToExcelWithOpenXML(WorkbookPart workbookPart, Sheets sheets, List<Transaction> transactions, string sheetPrefix)
        {
            var groupedTransactions = transactions.GroupBy(a => new Tuple<int, int>(a.Date.Month, a.Date.Year));

            foreach (var transaction in groupedTransactions)//iterate transactions by month and year
            {
                string sheetName = $"{transaction.Key.Item1} {transaction.Key.Item2}";
                var transactionByAccountNameGroup = transaction.ToList().GroupBy(a => a.Name); //groups of transaction by account name
                foreach (var accountGroup in transactionByAccountNameGroup)
                {
                    List<GroupedTransaction> accountTransactions = accountGroup.Select(a => new GroupedTransaction()
                    {
                        TotalAmount = a.Amount,
                        Date = a.Date.ToShortDateString(),
                        Name = a.Name
                    }).ToList();


                    //AddToSheet(workbookPart,sheets,sheetName, accountTransactions);
                }
            }
        }

        private void AddToSheet(
            WorkbookPart workbookPart,
            Sheets sheets,
            string sheetName,
            List<string> headerRow,
            Tuple<List<CellValues>,List<List<string>>> data,
            List<GroupedTransaction> groupedData)
        {
            int rowsCount = -1;
            Sheet sheet = null;
            SheetData sheetData = null;
            WorksheetPart worksheetPart = null;
            sheet = GetSheetByName(workbookPart, sheetName);
            if (sheet == null)//create new sheet
            {
                worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                sheet = new Sheet() { Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = (uint)(sheets.Count() + 1), Name = sheetName };
                sheets.Append(sheet);

                Worksheet worksheet = new Worksheet();
                sheetData = new SheetData();
                worksheet.Append(sheetData);
                worksheetPart.Worksheet = worksheet;
                AddRowToSheetData(sheetData, 1, headerRow, data.Item1);
                rowsCount = 2;
            }
            else
            {
                worksheetPart = GetWorksheetPartBySheet(workbookPart, sheet);
                sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();
                rowsCount = sheetData.Elements<Row>().Count();
            }
            
            foreach(var dataRowToAdd in data.Item2)
            {
                AddRowToSheetData(sheetData, (uint)rowsCount, dataRowToAdd, data.Item1);
                rowsCount++;
            }
        }
        private void AddRowToSheetData(SheetData sheetData, uint rowIndex, List<string> cellValues,List<CellValues> cellTypes)
        {
            Row row = new Row { RowIndex = rowIndex };

            for(int i = 0; i < cellValues.Count; i++)
            {
                string cellValue = cellValues[i];
                CellValues cellType = cellTypes[i];

                Cell cell = new Cell
                {
                    DataType = cellType,
                    CellValue = new CellValue(cellValue)
                };
                row.Append(cell);
            }
            sheetData.Append(row);
        }

        //private void AddToSheetWithSL(SLDocument sl, string sheetName, List<GroupedTransaction> groupedData)
        //{
        //    if (!IsSheetExist(sl,sheetName))
        //    {
        //        sl.AddWorksheet(sheetName);
        //        sl.SetCellValue(1, 2, "Name");
        //        sl.SetCellValue(1, 3, "Total Amount");
        //        sl.SetCellValue(1, 1, "Date");
        //    }

        //    sl.SelectWorksheet(sheetName);




        //    foreach (var data in groupedData)
        //    {
        //        sl.SetCellValue(row, 1, data.Name);
        //        sl.SetCellValue(row, 2, data.TotalAmount);
        //        sl.SetCellValue(row, 3, data.Date);
        //        row++;
        //    }
        //}

        //private bool IsSheetExist(SLDocument sl, string sheetName)
        //{
        //    List<string> currentSheetNames = sl.GetSheetNames();
        //    return currentSheetNames.Contains(sheetName);
        //}

        //public List<Transaction> ParseExpenses(string path)
        //{
        //    string redundantLineName = @"סה""כ";

        //    List<Transaction> transactions = new List<Transaction>();
        //    try
        //    {
        //        using (SLDocument sl = new SLDocument(path))
        //        {
        //            int iRow = 2;  // assuming first row has headers

        //            DateTime effectiveDate = DateTime.MinValue;
        //            while (!string.IsNullOrEmpty(sl.GetCellValueAsString(iRow, 3)))
        //            {

        //                string dateCellVal = sl.GetCellValueAsString(iRow, 2);

        //                if (string.IsNullOrEmpty(dateCellVal) ||
        //                    !string.IsNullOrEmpty(dateCellVal) && DateTime.TryParse(dateCellVal, out effectiveDate))
        //                {
        //                    string nameCellVal = sl.GetCellValueAsString(iRow, 3);
        //                    string amountCellVal = sl.GetCellValueAsString(iRow, 7);
        //                    if (nameCellVal != redundantLineName)
        //                    {
        //                        transactions.Add(new Transaction()
        //                        {
        //                            Name = nameCellVal,
        //                            Amount = decimal.Parse(amountCellVal),
        //                            Date = effectiveDate
        //                        });
        //                    }
        //                }



        //                iRow++;
        //            }
        //        }
        //    }
        //    catch (Exception e)
        //    {

        //    }


        //    return transactions;
        //}

        public List<Transaction> ParseExpenses(Stream fileStream)
        {
            string redundantLineName = @"סה""כ";
            
            Dictionary<string, List<List<string>>> retData = new Dictionary<string, List<List<string>>>();

            List<Transaction> transactions = new List<Transaction>();
            try
            {
                using (SpreadsheetDocument document = SpreadsheetDocument.Open(fileStream, isEditable: false))
                {
                    WorkbookPart workbookPart = document.WorkbookPart;

                    foreach (Sheet sheet in workbookPart.Workbook.Sheets.Elements<Sheet>())
                    {
                        List<List<string>> sheetData = new List<List<string>>();
                        WorksheetPart worksheetPart = GetWorksheetPartBySheet(workbookPart, sheet);

                        foreach(Row row in worksheetPart.Worksheet.GetFirstChild<SheetData>().Elements<Row>())
                        {
                            List<string> rowData = new List<string>();
                            foreach(Cell cell in row.Elements<Cell>())
                            {
                                string cellValue = GetCellValue(worksheetPart, cell);
                                rowData.Add(cellValue);
                            }
                        }
                    }
                }
            }
            catch (Exception e)
            {

            }


            return transactions;
        }

        private string GetCellValue(WorksheetPart worksheetPart, Cell cell)
        {
            string value = cell.CellValue.InnerText;
            if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
            {
                var stringTable = worksheetPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
                if (stringTable != null)
                {
                    value = stringTable.SharedStringTable.ElementAt(int.Parse(value)).InnerText;
                }
            }
            return value;
        }

        private bool SheetExists(WorkbookPart workbookPart, string sheetName)
        {
            // Ensure that the workbook part and sheet name are valid
            if (workbookPart == null || string.IsNullOrWhiteSpace(sheetName))
            {
                return false;
            }

            // Get the Sheets collection from the Workbook
            Sheets sheets = workbookPart.Workbook.Sheets;

            // Iterate through all sheets to find if any sheet matches the specified name
            foreach (Sheet sheet in sheets)
            {
                if (sheet.Name == sheetName)
                {
                    return true;
                }
            }

            // Return false if no sheet with the specified name is found
            return false;
        }
        private Sheet GetSheetByName(WorkbookPart workbookPart, string sheetName)
        {
            // Ensure that the workbook part and sheet name are valid
            if (workbookPart == null || string.IsNullOrWhiteSpace(sheetName))
            {
                return null;
            }

            // Get the Sheets collection from the Workbook
            Sheets sheets = workbookPart.Workbook.Sheets;

            // Find the sheet with the specified name
            foreach (Sheet sheet in sheets.OfType<Sheet>())
            {
                if (sheet.Name == sheetName)
                {
                    return sheet;
                }
            }

            return null;
        }
        private WorksheetPart GetWorksheetPartBySheet(WorkbookPart workbookPart, Sheet sheet)
        {
            if (sheet == null)
            {
                return null;
            }

            // Get the relationship ID of the sheet
            string relationshipId = sheet.Id;

            // Use the relationship ID to get the corresponding WorksheetPart
            WorksheetPart worksheetPart = (WorksheetPart)workbookPart.GetPartById(relationshipId);

            return worksheetPart;
        }

        private string GetSaveFilePath()
        {
            var saveFileDialog = new Microsoft.Win32.SaveFileDialog
            {
                Filter = "Excel files (*.xlsx)|*.xlsx",
                FileName = "Output.xlsx"
            };
            return saveFileDialog.ShowDialog() == true ? saveFileDialog.FileName : throw new Exception("path for save not provided") ;
        }
    }
}
