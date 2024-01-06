using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CashFlowApp.UI.model
{
    public class ExcelBuilderModel
    {
        public List<Transaction> Expenses { get; set; }
        public List<Transaction> StandingOrders { get; set; }
        public List<Transaction> Incomes { get; set; }
    }
}
