using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GameHaven
{
    public class MagicCard
    {
        public int UnitOfMeasure { get; set; }
        public string Name { get; set; }
        public string Attribute { get; set; }
        public string Expansion { get; set; }
        public string Rarity { get; set; }
        public string Color { get; set; }
        public int Quantity { get; set; }
        public double Price { get; set; }
        public double Factor
        {
            get
            {
                switch (this.UnitOfMeasure)
                {
                    case 1:
                        if (this.Quantity <= 8)
                            return 0.7;
                        if (this.Quantity <= 16)
                            return 0.5;
                        return 0.25; // >= 17
                    case 2:
                        if (this.Quantity <= 4)
                            return 0.7;
                        if (this.Quantity <= 8)
                            return 0.5;
                        if (this.Quantity <= 16)
                            return 0.33;
                        if (this.Quantity <= 20)
                            return 0.25;
                        return 0; // >= 21
                    case 3:
                        if (this.Quantity <= 4)
                            return 0.5;
                        if (this.Quantity <= 8)
                            return 0.33;
                        if (this.Quantity <= 12)
                            return 0.25;
                        return 0; // >= 13
                    case 4:
                        if (this.Quantity <= 4)
                            return 0.5;
                        if (this.Quantity <= 8)
                            return 0.25;
                        return 0; // >= 9
                    default:
                        return 0;
                }
            }
        }
        public double PayoutCredit
        {
            get { return Math.Round(this.Price * this.Factor, 2); }
        }
        public double CalculatePayoutCash(double cashFactor)
        {
            return Math.Round(this.PayoutCredit * cashFactor, 2);
        }
    }
}
