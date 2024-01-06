using CashFlowApp.UI.model;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CashFlowApp.UI
{
    public class CashFlowFacade
    {
        internal void GenerateOutput(CashFlowFilesModel cashFlowFilesModel)
        {
            ExcelBuilder excelBuilder = new ExcelBuilder();
            ExcelBuilderModel model = new ExcelBuilderModel();
            TransactionProcessor processor = new TransactionProcessor();

            //List<Transaction> expenses = excelBuilder.ParseExpenses(cashFlowFilesModel.ExpensesFilePath);
            //List<Transaction> incomes = excelBuilder.ParseFile(IncomesFilePath.Text);
            //model.Expenses = expenses;
            excelBuilder.CreateExcel(model);
        }
    }
}
