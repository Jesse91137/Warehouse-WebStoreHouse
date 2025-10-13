using System.Collections.Generic;

namespace WebStoreHouse.Models
{
    /// <summary>
    /// 表示一個電子郵件工作項目，包含寄件人、收件人、主旨、內容及SMTP相關設定。
    /// </summary>
    public class EmailWorkItem
    {
        /// <summary>
        /// 寄件者電子郵件地址。
        /// </summary>
        public string From { get; set; }

        /// <summary>
        /// 收件者電子郵件地址清單。
        /// </summary>
        public List<string> To { get; set; }

        /// <summary>
        /// 副本收件者 (CC) 電子郵件地址清單。
        /// </summary>
        public List<string> Cc { get; set; }

        /// <summary>
        /// 郵件主旨。
        /// </summary>
        public string Subject { get; set; }

        /// <summary>
        /// 郵件內容。
        /// </summary>
        public string Body { get; set; }

        /// <summary>
        /// 指定郵件內容是否為HTML格式，預設為true。
        /// </summary>
        public bool IsBodyHtml { get; set; } = true;

        /// <summary>
        /// SMTP伺服器位址。
        /// </summary>
        public string SmtpServer { get; set; }

        /// <summary>
        /// SMTP伺服器連接埠，預設為25。
        /// </summary>
        public int SmtpPort { get; set; } = 25;

        /// <summary>
        /// SMTP使用者名稱。
        /// </summary>
        public string SmtpUser { get; set; }

        /// <summary>
        /// SMTP密碼。
        /// </summary>
        public string SmtpPass { get; set; }

        /// <summary>
        /// 郵件傳送逾時時間（毫秒），預設為15000。
        /// </summary>
        public int TimeoutMs { get; set; } = 15000;
    }
}