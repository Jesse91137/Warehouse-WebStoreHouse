using System.Collections.Generic; // 引入泛型集合命名空間，提供 IEnumerable<T> 型別
using System.Threading.Tasks; // 引入非同步程式設計相關命名空間，提供 Task 型別

namespace WebStoreHouse.Services // 定義 WebStoreHouse.Services 命名空間
{
    /// <summary>
    /// 簡單的 Email Notification 介面，方便測試與替換實作
    /// 新增 20251008 by Jesse
    /// </summary>
    public interface IEmailNotificationService // 定義 Email 通知服務介面
    {
        /// <summary>
        /// 取得指定 dropship 物件的收件人清單
        /// </summary>
        /// <param name="dropship">動態型別的 dropship 物件</param>
        /// <returns>收件人電子郵件字串集合</returns>
        IEnumerable<string> GetRecipientsForDropship(dynamic dropship); // 宣告取得收件人方法

        /// <summary>
        /// 非同步發送 dropship 編輯通知郵件
        /// </summary>
        /// <param name="dropship">動態型別的 dropship 物件</param>
        /// <returns>非同步作業 Task</returns>
        Task SendDropshipEditedEmailAsync(dynamic dropship); // 宣告發送編輯通知郵件方法

        /// <summary>
        /// 非同步發送 dropship 刪除通知郵件
        /// </summary>
        /// <param name="dropship">動態型別的 dropship 物件</param>
        /// <returns>非同步作業 Task</returns>
        Task SendDropshipDeletedEmailAsync(dynamic dropship); // 宣告發送刪除通知郵件方法
    }
}
