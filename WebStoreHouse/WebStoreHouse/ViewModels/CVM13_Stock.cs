using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using WebStoreHouse.Models;

namespace WebStoreHouse.ViewModels
{
    public class CVM13_Stock
    {
        public CVM13_Stock()
        {
            storeHouseStock_SCs = new List<E_StoreHouseStock_SC>();
            storeHouseStocks = new List<E_StoreHouseStock>();
        }
        
        public List<E_StoreHouseStock_SC> storeHouseStock_SCs { get; set; }
        public List<E_StoreHouseStock> storeHouseStocks { get; set; }        
        //public List<E_ZRSD13> ZRSD13s { get; set; }
    }
}