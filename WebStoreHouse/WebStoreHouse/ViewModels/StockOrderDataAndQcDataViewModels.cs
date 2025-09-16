using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

using System.Data;
using WebStoreHouse.Models;

namespace WebStoreHouse.ViewModels
{
    public class StockOrderDataAndQcDataViewModels
    {
        public DataSet StockData { get; set; }
        public List<QC_Model.QCData> QCData { get; set; }
    }
}