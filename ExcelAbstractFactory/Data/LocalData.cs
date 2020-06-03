using ExcelAbstractFactory.Models;
using System;
using System.Collections.Generic;

namespace ExcelAbstractFactory.Data
{
    public static class LocalData
    {
        public static List<Product> GetProductList()
        {
            return new List<Product>()
            {
                new Product { ProductID = 1, ProductName = "Bath Rug", CategoryID = 1, UnitPrice = 24.5m, OutOfStock = false, StockDate = Convert.ToDateTime("2013-07-10")},
                new Product { ProductID = 2, ProductName = "Shower Curtain", CategoryID = 1, UnitPrice = 30.99m, OutOfStock = false, StockDate = Convert.ToDateTime("2013-07-13")},
                new Product { ProductID = 3, ProductName = "Soap Dispenser", CategoryID = 1, UnitPrice = 12.4m, OutOfStock = true, StockDate = null},
                new Product { ProductID = 4, ProductName = "Toilet Tissue", CategoryID = 1, UnitPrice = 15, OutOfStock = false, StockDate = Convert.ToDateTime("2013-05-16")},
                new Product { ProductID = 5, ProductName = "Branket", CategoryID = 2, UnitPrice = 60, OutOfStock = false, StockDate = Convert.ToDateTime("2013-08-22")},
                new Product { ProductID = 6, ProductName = "Mattress Protector", CategoryID = 2, UnitPrice = 30.4m, OutOfStock = true, StockDate = null },
                new Product { ProductID = 8, ProductName = "Baking Pan", CategoryID = 3, UnitPrice = 10.99m, OutOfStock = false, StockDate = Convert.ToDateTime("2013-10-26")},
                new Product { ProductID = 9, ProductName = "Coffee Maker", CategoryID = 3, UnitPrice = 49.39m, OutOfStock = false, StockDate = Convert.ToDateTime("2013-09-10")},
                new Product { ProductID = 11, ProductName = "Pressure Cooker", CategoryID = 3, UnitPrice = 90.5m, OutOfStock = true, StockDate = null},
                new Product { ProductID = 12, ProductName = "Water Pitcher", CategoryID = 3, UnitPrice = 29.99m, OutOfStock = false, StockDate = Convert.ToDateTime("2013-12-08")}			
            };
        }

        public static List<Product> GetProductList2()
        {
            return new List<Product>()
            {
                new Product { ProductID = 111, ProductName = "Elma", CategoryID = 1, UnitPrice = 24.5m, OutOfStock = false, StockDate = Convert.ToDateTime("2013-07-10")},
                new Product { ProductID = 112, ProductName = "Armut", CategoryID = 1, UnitPrice = 30.99m, OutOfStock = false, StockDate = Convert.ToDateTime("2013-07-13")},
                new Product { ProductID = 113, ProductName = "Kiraz", CategoryID = 1, UnitPrice = 12.4m, OutOfStock = true, StockDate = null},
                new Product { ProductID = 114, ProductName = "Portakal", CategoryID = 1, UnitPrice = 15, OutOfStock = false, StockDate = Convert.ToDateTime("2013-05-16")},
                new Product { ProductID = 115, ProductName = "Kayısı", CategoryID = 2, UnitPrice = 60, OutOfStock = false, StockDate = Convert.ToDateTime("2013-08-22")}
            };
        }
    }
}
