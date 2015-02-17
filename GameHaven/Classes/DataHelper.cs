using System;
using System.Diagnostics;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using System.Windows.Forms;

namespace GameHaven
{
    public class DataHelper
    {
        const int THRESHOLD = 2;

        public static string GetStartingPoint(Excel.Range dataRange)
        {
            int colsWithData;
            string dataStartAddr = string.Empty;
            foreach (Excel.Range row in dataRange.Rows)
            {
                colsWithData = 0;
                foreach (Excel.Range col in row.Columns)
                {
                    if (col.Value2 != null) colsWithData++;
                    if (colsWithData > THRESHOLD)
                    {
                        dataStartAddr = row.Address;
                        break;
                    }
                }
                if (dataStartAddr != string.Empty) break;
            }
            return dataStartAddr;
        }

        public static Excel.Range GetColumnHeaderLocations(string dataStartAddr,
            Excel.Range dataRange)
        {
            string firstHeader = dataStartAddr.Split(':')[0];
            string lastHeader = dataStartAddr.Split(':')[1];
            string headerText = string.Empty;
            Excel.Range headerRange = dataRange.Range[firstHeader, lastHeader];
            foreach (Excel.Range cell in headerRange)
            {
                if (cell.Value2 != null)
                {
                    headerText = (string)cell.Value2.ToLower();
                    if (headerText.Contains(ColumnHeaderNames.UnitOfMeasure))
                    {
                        ColumnHeaderAddresses.UnitOfMeasure = cell.Address;
                        continue;
                    }
                    if (headerText.Contains(ColumnHeaderNames.Attribute))
                    {
                        ColumnHeaderAddresses.Attribute = cell.Address;
                        continue;
                    }
                    if (headerText.Contains(ColumnHeaderNames.Name))
                    {
                        ColumnHeaderAddresses.Name = cell.Address;
                        continue;
                    }
                    if (headerText.Contains(ColumnHeaderNames.Expansion))
                    {
                        ColumnHeaderAddresses.Expansion = cell.Address;
                        continue;
                    }
                    if (headerText.Contains(ColumnHeaderNames.Rarity))
                    {
                        ColumnHeaderAddresses.Rarity = cell.Address;
                        continue;
                    }
                    if (headerText.Contains(ColumnHeaderNames.Color))
                    {
                        ColumnHeaderAddresses.Color = cell.Address;
                        continue;
                    }
                    if (headerText.Contains(ColumnHeaderNames.Price))
                    {
                        ColumnHeaderAddresses.Price = cell.Address;
                        continue;
                    }
                    if (headerText.Contains(ColumnHeaderNames.Quantity))
                    {
                        ColumnHeaderAddresses.Quantity = cell.Address;
                        continue;
                    }
                }
            }
            return headerRange;
        }

        public static List<MagicCard> GetMagicCardsFromData(Excel.Range dataRange,
            Excel.Range headerRange, ProgressBar bar)
        {
            if (bar != null)
            {
                bar.Minimum = 1;
                bar.Maximum = dataRange.Rows.Count;
                bar.Value = 1;
                bar.Step = 1;
            }
            List<MagicCard> cards = new List<MagicCard>();
            Debug.WriteLine(dataRange.Address);
            int headerRow = Convert.ToInt32(
                headerRange.Address.Substring(headerRange.Address.Length - 1, 1)
            );
            string firstDataCell = string.Format("$A${0}", ++headerRow);
            string lastDataCell = dataRange.Address.Split(':')[1];
            Debug.WriteLine(firstDataCell);
            Excel.Range mainRange = dataRange.Range[firstDataCell, lastDataCell];
            foreach (Excel.Range row in mainRange.Rows)
            {
                int uom = 0;
                string name = string.Empty;
                string attr = string.Empty;
                string rarity = string.Empty;
                string color = string.Empty;
                string expansion = string.Empty;
                int qty = 0;
                double price = 0;
                foreach (Excel.Range col in row.Columns)
                {
                    if (col.Value2 == null) continue;
                    try
                    {
                        var val = col.Value2;
                        string addr = col.Address.Substring(0, 2);
                        if (addr == ColumnHeaderAddresses.UnitOfMeasure.Substring(0, 2))
                        {
                            uom = Convert.ToInt32(val);
                            continue;
                        }
                        if (addr == ColumnHeaderAddresses.Name.Substring(0, 2))
                        {
                            name = col.Value2;
                            continue;
                        }
                        if (addr == ColumnHeaderAddresses.Attribute.Substring(0, 2))
                        {
                            attr = col.Value2;
                            continue;
                        }
                        if (addr == ColumnHeaderAddresses.Rarity.Substring(0, 2))
                        {
                            rarity = col.Value2;
                            continue;
                        }
                        if (addr == ColumnHeaderAddresses.Color.Substring(0, 2))
                        {
                            color = col.Value2;
                            continue;
                        }
                        if (addr == ColumnHeaderAddresses.Expansion.Substring(0, 2))
                        {
                            expansion = col.Value2;
                            continue;
                        }
                        if (addr == ColumnHeaderAddresses.Quantity.Substring(0, 2))
                        {
                            qty = Convert.ToInt32(col.Value2);
                            continue;
                        }
                        if (addr == ColumnHeaderAddresses.Price.Substring(0, 2))
                        {
                            price = Convert.ToDouble(col.Value2);
                            continue;
                        }
                    }
                    catch (Exception e)
                    {
                        Debug.WriteLine(e.ToString());
                    }
                }
                cards.Add(new MagicCard()
                {
                    UnitOfMeasure = uom,
                    Name = name,
                    Attribute = attr,
                    Expansion = expansion,
                    Color = color,
                    Rarity = rarity,
                    Quantity = qty,
                    Price = price
                });
                if (bar != null) bar.PerformStep();
            }
            return cards;
        }

        public static List<MagicCard> GroupCardsByName(List<MagicCard> cards)
        {
            var groupedCards = cards.GroupBy(c => c.Name);
            List<MagicCard> cardTotals = new List<MagicCard>();
            foreach (var cardGroup in groupedCards)
            {
                List<MagicCard> cardsByName = cardGroup.ToList();
                int uom = 0;
                string name = string.Empty;
                string attr = string.Empty;
                string rarity = string.Empty;
                string color = string.Empty;
                string expansion = string.Empty;
                int qty = 0;
                double price = 0;

                foreach (MagicCard card in cardsByName)
                {
                    uom = card.UnitOfMeasure;
                    name = card.Name;
                    expansion = card.Expansion;
                    rarity = card.Rarity;
                    color = card.Color;
                    price = card.Price;
                    qty += card.Quantity;
                }
                cardTotals.Add(new MagicCard()
                {
                    UnitOfMeasure = uom,
                    Name = name,
                    Expansion = expansion,
                    Rarity = rarity,
                    Color = color,
                    Price = price,
                    Quantity = qty
                });
            }
            return cardTotals;
        }

        public static void WriteToHtmlFile(FileStream fileStream, List<MagicCard> cards, 
            double factor)
        {
            StringBuilder htmlString = new StringBuilder();
            htmlString.Append(
                @"<!doctype html><html><head><title>Buy List</title><style> body { font-family: Arial; } table { width: 90%; margin: 0 auto; } 
                th { text-align:left; } td { font-size: 0.8em; }
                </style>
                </head><body><center><h1>BUY LIST</h1>
                <h5>Prices are based on near mint fair market price and are therefore subject to change without notice.</h5>");
            htmlString.Append(
                String.Format("<i>Last updated {0}</i>",
                    DateTime.Today.ToShortDateString()));
            var expansions = cards.Select(c => c.Expansion).Distinct();
            foreach (var expansion in expansions)
            {
                htmlString.Append(String.Format("<a href=\"#{0}\">{1}</a><br>", expansion.Replace(" ", ""), expansion));
            }
            htmlString.Append(
                @"<h2>Magic: the Gathering</h2></center><table><tr><th>Name</th><th>Expansion</th><th>Rarity</th><th>Color</th>
                <th>Payout Credit</th><th>Payout Cash</th></tr>");

            // Show all card data in the table
            string currentExpansion = string.Empty;
            string expansionText = string.Empty;
            foreach (MagicCard card in cards)
            {
                if (card.Expansion != currentExpansion)
                {
                    currentExpansion = card.Expansion;
                    expansionText = String.Format(
                        "<span id=\"{0}\">{1}</span>", 
                        card.Expansion.Replace(" ", ""), 
                        card.Expansion);
                }
                else
                {
                    expansionText = String.Format("<span>{0}</span>", card.Expansion);
                }
                htmlString.Append(
                    String.Format(
                        "<tr><td>{0}</td><td>{1}</td><td>{2}</td><td>{3}</td><td>{4:C2}</td><td>{5:C2}</td></tr>",
                        card.Name,
                        expansionText,
                        card.Rarity,
                        card.Color,
                        card.PayoutCredit,
                        card.CalculatePayoutCash(factor)));
            }

            htmlString.Append("</table></body></html>");
            using (StreamWriter file = new StreamWriter(fileStream))
            {
                file.Write(htmlString.ToString());
            }
        }
    }
}
