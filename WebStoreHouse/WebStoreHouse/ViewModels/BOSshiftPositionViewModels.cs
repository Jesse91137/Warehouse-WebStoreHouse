using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using WebStoreHouse.Models;

namespace WebStoreHouse.ViewModels
{
    public class BOSshiftPositionViewModels
    {
        public BOSshiftPositionViewModels()
        {
            storeHouseStock_BOSs = new List<E_StoreHouseStock_BOS>();
            storeHouseStocks = new List<E_StoreHouseStock>();
        }

        public List<E_StoreHouseStock_BOS> storeHouseStock_BOSs { get; set; }
        public List<E_StoreHouseStock> storeHouseStocks { get; set; }
    }
}