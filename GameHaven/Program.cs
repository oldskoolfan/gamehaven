using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Diagnostics;
using Excel = Microsoft.Office.Interop.Excel;

namespace GameHaven
{
    class Program
    {
        const double CASH_FACTOR = 0.6;
        static void Main(string[] args)
        {
            Excel.Application app = new Excel.Application();
            Excel.Workbook book = app.Workbooks.Open(@"C:\MyProjects\GameHaven\testdata3.xlsx");
            Excel.Worksheet sheet = book.ActiveSheet;
            Excel.Range dataRange = sheet.UsedRange;

            // get our data starting point
            string dataStartAddr = DataHelper.GetStartingPoint(dataRange);
            Debug.WriteLine(String.Format("Data starts here: {0}", dataStartAddr));

            // get our column headers and locations           
            Excel.Range headerRange = DataHelper.GetColumnHeaderLocations(
                dataStartAddr, 
                dataRange);

            // main loop through data to get card info
            List<MagicCard> cards = DataHelper.GetMagicCardsFromData(
                dataRange,
                headerRange,
                null);

            // since we have a lot of double rows, group by name
            List<MagicCard> cardTotals = DataHelper.GroupCardsByName(cards);

            // finally we can output results
            FileStream stream = new FileStream(
                string.Format(
                    @"C:\MyProjects\GameHaven\testfile_{0}{1}{2}.html",
                    DateTime.Today.Year,
                    DateTime.Today.Month,
                    DateTime.Today.Day), 
                FileMode.OpenOrCreate);
            DataHelper.WriteToHtmlFile(stream, cardTotals, CASH_FACTOR);

            //Console.Read();
            book.Close();
            app.Quit();
            return;
        }
    }
}
