using System;

namespace ExcelAbstractFactory.Models
{
    public class Product
    {
        public int ProductID { get; set; }
        public string ProductName { get; set; }
        public int? CategoryID { get; set; }
        public decimal? UnitPrice { get; set; }
        public bool OutOfStock { get; set; }
        public DateTime? StockDate { get; set; }
    }
}
