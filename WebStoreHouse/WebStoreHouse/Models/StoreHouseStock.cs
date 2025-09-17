using System;
using System.Data.Entity;
using System.Linq;

namespace WebStoreHouse.Models
{
    public partial class StoreHouseStock
    {
        public int serialno { get; set; }
        /// <summary>
        /// 工單
        /// </summary>
        public string nowono { get; set; }
        public Nullable<int> sno { get; set; }
        public string wono { get; set; }
        /// <summary>
        /// 客戶業單
        /// </summary>
        public string cust_wono { get; set; }
        /// <summary>
        /// 機種名稱
        /// </summary>
        public string eng_sr { get; set; }
        /// <summary>
        /// 訂單數量
        /// </summary>
        public int order_count { get; set; }
        /// <summary>
        /// 數量
        /// </summary>
        public int quantity { get; set; }
        /// <summary>
        /// 箱數
        /// </summary>
        public int box_quantity { get; set; }
        /// <summary>
        /// KF10
        /// </summary>
        public int kf10 { get; set; }
        /// <summary>
        /// KQ30
        /// </summary>
        public int kq30 { get; set; }
        public string sap_in { get; set; }
        /// <summary>
        /// 儲位
        /// </summary>
        public string position { get; set; }
        public int acc_in { get; set; }
        public int outed { get; set; }
        public int notout { get; set; }
        public int borrow { get; set; }
        /// <summary>
        /// 預計銷單日
        /// </summary>
        public Nullable<System.DateTime> due_date { get; set; }
        /// <summary>
        /// 備註
        /// </summary>
        public string mark { get; set; }
        /// <summary>
        /// 入庫日期
        /// </summary>
        public string inputdate { get; set; }
        /// <summary>
        /// 包裝
        /// </summary>
        public string package { get; set; }
        public string output_local { get; set; }
        public string Igroup { get; set; }
    }
}