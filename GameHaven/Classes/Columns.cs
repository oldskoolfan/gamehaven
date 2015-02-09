using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GameHaven
{
    class ColumnHeaderNames
    {
        public static readonly string UnitOfMeasure = "measure";
        public static readonly string Name = "name";
        public static readonly string Attribute = "attribute";
        public static readonly string Expansion = "size";
        public static readonly string Rarity = "rarity";
        public static readonly string Color = "color";
        public static readonly string Quantity = "qty";
        public static readonly string Price = "price";
    }

    class ColumnHeaderAddresses
    {
        public static string UnitOfMeasure { get; set; }
        public static string Name { get; set; }
        public static string Attribute { get; set; }
        public static string Expansion { get; set; }
        public static string Rarity { get; set; }
        public static string Color { get; set; }
        public static string Quantity { get; set; }
        public static string Price { get; set; }
    }
}
