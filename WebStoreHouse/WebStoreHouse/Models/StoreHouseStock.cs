using System;
using System.Data.Entity;
using System.Linq;

namespace WebStoreHouse.Models
{
    public partial class StoreHouseStock
    {
        public int serialno { get; set; }
        public string nowono { get; set; }
        public Nullable<int> sno { get; set; }
        public string wono { get; set; }
        public string cust_wono { get; set; }
        public string eng_sr { get; set; }
        public int order_count { get; set; }
        public int quantity { get; set; }
        public int box_quantity { get; set; }
        public int kf10 { get; set; }
        public int kq30 { get; set; }
        public string sap_in { get; set; }
        public string position { get; set; }
        public int acc_in { get; set; }
        public int outed { get; set; }
        public int notout { get; set; }
        public int borrow { get; set; }
        public Nullable<System.DateTime> due_date { get; set; }
        public string mark { get; set; }
        public string inputdate { get; set; }
        public string package { get; set; }
        public string output_local { get; set; }
        public string Igroup { get; set; }
    }
}