using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using WebStoreHouse.Models;


namespace WebStoreHouse.ViewModels
{
    public class StockSCViewModels
    {      
        public List<E_StoreHouseStock_SC> Stock_SCs { get; set; }
        public List<E_StoreHouseStock_BOS> stock_BOSs { get; set; }
        public int CurrentPage { get; set; }
        public int TotalPages { get; set; }
    }
}