using System;

namespace WebStoreHouse.Models
{
    /// <summary>
    /// 直出資料-查詢資料的資料傳輸物件。
    /// </summary>
    public class EDropshippingQDto
    {
        /// <summary>
        /// 序號。
        /// </summary>
        public int sno { get; set; }

        /// <summary>
        /// 出貨日期（字串格式）。
        /// </summary>
        public string date { get; set; }

        /// <summary>
        /// 送貨單號。
        /// </summary>
        public string DN { get; set; }

        /// <summary>
        /// 機種編號。
        /// </summary>
        public string eng_sr { get; set; }

        /// <summary>
        /// 數量。
        /// </summary>
        public Nullable<int> quantity { get; set; }

        /// <summary>
        /// 貨運行資訊。
        /// </summary>
        public string freight { get; set; }

        /// <summary>
        /// 是否已檢查完成。
        /// </summary>
        public Nullable<bool> checkOK { get; set; }

        /// <summary>
        /// 工單號碼。
        /// </summary>
        public string wono { get; set; }

        /// <summary>
        /// 滿箱數量（可為 null）。
        /// </summary>
        public int? Full_Amount { get; set; }

        /// <summary>
        /// 尾數箱數量。
        /// </summary>
        /// <value>
        /// 數量 / 滿箱數量 餘數。
        /// </value>
        public string NumberOfTailBoxes { get; set; }
    }
}