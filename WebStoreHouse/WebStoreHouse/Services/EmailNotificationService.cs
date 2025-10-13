using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Web;
using WebStoreHouse.Models;

namespace WebStoreHouse.Services
{
    /// <summary>
    /// Email 寄送相關邏輯集中在此類別，方便測試與重用
    /// 新增 20251008 by Jesse
    /// </summary>
    public class EmailNotificationService : IEmailNotificationService
    {
        /// <summary>
        /// 建構函式，初始化 EmailNotificationService 實例
        /// </summary>
        public EmailNotificationService()
        {
            // 讀取 SMTP 設定
            ReadSmtpSettings(out smtpServer, out smtpPort, out smtpUser, out smtpPass);
        }

        #region 共用變數
        // WHY: 共用變數，避免重複讀取設定
        /// <summary>
        /// SMTP 伺服器位址。
        /// </summary>
        string smtpServer = null;
        /// <summary>
        /// SMTP 使用者名稱。
        /// </summary>
        string smtpUser = null;
        /// <summary>
        /// SMTP 密碼。
        /// </summary>
        string smtpPass = null;
        /// <summary>
        /// SMTP 連接埠，預設為 25。
        /// </summary>
        int smtpPort = 25;
        /// <summary>
        /// 收件者電子郵件地址集合。
        /// </summary>
        IEnumerable<string> recipientsList = null;
        /// <summary>
        /// 副本收件者（CC）電子郵件地址集合。
        /// </summary>
        IEnumerable<string> ccList = null;
        /// <summary>
        /// 寄件者電子郵件地址。
        /// </summary>
        string fromAddress = null;

        /// <summary>
        /// Email 發送工作項目。用於暫存本次欲寄送的 Email 內容與設定。
        /// </summary>
        EmailWorkItem emailWork = null;
        #endregion

        #region 寄信相關共用方法
        /// <summary>
        /// 讀取 SMTP 設定，並回傳相關參數。
        /// </summary>
        /// <param name="smtpServer">SMTP 伺服器位址。</param>
        /// <param name="smtpPort">SMTP 伺服器連接埠，預設為 25。</param>
        /// <param name="smtpUser">SMTP 使用者名稱。</param>
        /// <param name="smtpPass">SMTP 密碼。</param>
        /// <summary>
        /// 讀取 SMTP 設定，並回傳相關參數。
        /// </summary>
        private static void ReadSmtpSettings(out string smtpServer, out int smtpPort, out string smtpUser, out string smtpPass)
        {
            smtpServer = System.Configuration.ConfigurationManager.AppSettings["SmtpServer"];
            smtpUser = System.Configuration.ConfigurationManager.AppSettings["SmtpUser"];
            smtpPass = System.Configuration.ConfigurationManager.AppSettings["SmtpPass"];
            smtpPort = int.TryParse(System.Configuration.ConfigurationManager.AppSettings["SmtpPort"], out var port) ? port : 25;
        }

        /// <summary>
        /// 建立一個 <see cref="EmailWorkItem"/> 實例，包含寄件人、收件人、主旨、內容及 SMTP 設定。
        /// </summary>
        /// <param name="to">收件者電子郵件地址集合。</param>
        /// <param name="from">寄件者電子郵件地址。</param>
        /// <param name="subject">郵件主旨。</param>
        /// <param name="body">郵件內容（HTML格式）。</param>
        /// <param name="smtpServer">SMTP 伺服器位址。</param>
        /// <param name="smtpPort">SMTP 伺服器連接埠。</param>
        /// <param name="smtpUser">SMTP 使用者名稱。</param>
        /// <param name="smtpPass">SMTP 密碼。</param>
        /// <param name="timeoutMs">郵件傳送逾時時間（毫秒），預設為 15000。</param>
        /// <returns>已初始化的 <see cref="EmailWorkItem"/> 物件。</returns>
        private static EmailWorkItem CreateWorkItem(IEnumerable<string> to, IEnumerable<string> cc, string from, string subject, string body, string smtpServer, int smtpPort, string smtpUser, string smtpPass, int timeoutMs = 15000)
        {
            // WHY: 集中初始化 EmailWorkItem，確保所有欄位正確
            return new EmailWorkItem
            {
                From = from,
                To = to?.ToList() ?? new List<string>(),
                Cc = cc?.ToList() ?? new List<string>(),
                Subject = subject,
                Body = body,
                IsBodyHtml = true,
                SmtpServer = smtpServer,
                SmtpPort = smtpPort,
                SmtpUser = smtpUser,
                SmtpPass = smtpPass,
                TimeoutMs = timeoutMs
            };
        }
        #endregion

        /// <summary>
        /// 寄送預計直出資料被編輯的通知 Email
        /// </summary>
        /// <param name="dropship">直出資料物件</param>
        /// <returns>非同步 Task</returns>
        public virtual async Task SendDropshipEditedEmailAsync(dynamic dropship)
        {

            // 檢查 dropship 物件是否為 null，若為 null 則直接結束方法以避免後續錯誤。
            if (dropship is null) return;

            /// <summary>
            /// 郵件主旨。
            /// </summary>
            string subject = "編輯：成倉系統內容修改請確認";

            try
            {
                // 取得收件者清單
                recipientsList = (GetRecipientsForDropship(dropship) as IEnumerable<string>)?.ToList() ?? new List<string>();
                if (!recipientsList.Any()) return; // WHY: 沒有收件者就不寄信
                // 取得副本收件者（CC）清單
                ccList = GetCcForDropship(dropship) ?? Enumerable.Empty<string>();
                // 寄件者電子郵件地址
                fromAddress = !string.IsNullOrWhiteSpace(smtpUser) ? smtpUser : "no-reply@eversun.com.tw";

                // WHY: 集中欄位轉換與 HTML 編碼，避免 XSS
                var bodySb = new System.Text.StringBuilder(512)
                    .AppendLine("<p>系統通知：以下預計直出資料已被修改，請確認內容是否正確。</p>")
                    .AppendLine("<table border='1' cellpadding='6' cellspacing='0'>")
                    .AppendLine("<tr><th>欄位</th><th>值</th></tr>")
                    .AppendLine($"<tr><td>Key (sno)</td><td>{ToHtmlString(dropship.sno)}</td></tr>")
                    .AppendLine($"<tr><td>工單號碼 (wono)</td><td>{ToHtmlString(dropship.wono)}</td></tr>")
                    .AppendLine($"<tr><td>日期</td><td>{ToHtmlString(dropship.date)}</td></tr>")
                    .AppendLine($"<tr><td>DN</td><td>{ToHtmlString(dropship.DN)}</td></tr>")
                    .AppendLine($"<tr><td>機種名稱 (eng_sr)</td><td>{ToHtmlString(dropship.eng_sr)}</td></tr>")
                    .AppendLine($"<tr><td>數量</td><td>{ToHtmlString(dropship.quantity)}</td></tr>")
                    .AppendLine($"<tr><td>貨運行</td><td>{ToHtmlString(dropship.freight)}</td></tr>")
                    .AppendLine("</table>")
                    .AppendLine("<p>如需更改或有疑問請聯絡成倉管理者。</p>");

                emailWork = CreateWorkItem(recipientsList, ccList, fromAddress, subject, bodySb.ToString(), smtpServer, smtpPort, smtpUser, smtpPass);
                WebStoreHouse.Services.BackgroundEmailQueue.Instance.Enqueue(emailWork);
                await Task.Yield();
            }
            catch (Exception ex)
            {
                LogEmailException(ex, nameof(SendDropshipEditedEmailAsync), smtpServer, smtpPort, smtpUser, fromAddress, recipientsList, ccList, subject);
                throw;
            }
        }

        /// <summary>
        /// 寄送預計直出資料被刪除的通知 Email
        /// 此方法會將刪除前的 預計直出 資料內容以 Email 通知相關人員，並記錄發送過程中發生的例外狀況。
        /// </summary>
        /// <param name="dropship">預計直出資料物件</param>
        /// <returns>非同步 Task，表示 Email 發送作業。</returns>
        /// <remarks>
        /// 1. 若 dropship 為 null 則不執行任何動作。
        /// 2. 會自動讀取 SMTP 設定、收件人、CC、寄件者等資訊。
        /// 3. 若收件人清單為空則不寄信。
        /// 4. 發送失敗時會將詳細例外資訊記錄至 Log/email_errors_yyyyMMdd.log。
        /// </remarks>
        public virtual async Task SendDropshipDeletedEmailAsync(dynamic dropship)
        {
            // 檢查 dropship 物件是否為 null，若為 null 則直接結束方法以避免後續錯誤。
            if (dropship is null) return;

            /// <summary>
            /// 郵件主旨。
            /// </summary>
            string subject = "刪除：成倉系統此筆資料請留意";

            try
            {
                // 取得收件者清單
                recipientsList = (GetRecipientsForDropship(dropship) as IEnumerable<string>)?.ToList() ?? new List<string>();
                // WHY: 沒有收件者就不寄信
                if (!recipientsList.Any()) return;
                // 取得副本收件者（CC）清單
                ccList = GetCcForDropship(dropship) ?? Enumerable.Empty<string>();
                // 寄件者電子郵件地址
                fromAddress = !string.IsNullOrWhiteSpace(smtpUser) ? smtpUser : "no-reply@eversun.com.tw";

                // WHY: 集中欄位轉換與 HTML 編碼，避免 XSS
                var bodySb = new System.Text.StringBuilder(512)
                    .AppendLine("<p>系統通知：以下為刪除前的預計直出資料，請留意。</p>")
                    .AppendLine("<table border='1' cellpadding='6' cellspacing='0'>")
                    .AppendLine("<tr><th>欄位</th><th>值</th></tr>")
                    .AppendLine($"<tr><td>Key (sno)</td><td>{ToHtmlString(dropship.sno)}</td></tr>")
                    .AppendLine($"<tr><td>工單號碼 (wono)</td><td>{ToHtmlString(dropship.wono)}</td></tr>")
                    .AppendLine($"<tr><td>日期</td><td>{ToHtmlString(dropship.date)}</td></tr>")
                    .AppendLine($"<tr><td>DN</td><td>{ToHtmlString(dropship.DN)}</td></tr>")
                    .AppendLine($"<tr><td>機種名稱 (eng_sr)</td><td>{ToHtmlString(dropship.eng_sr)}</td></tr>")
                    .AppendLine($"<tr><td>數量</td><td>{ToHtmlString(dropship.quantity)}</td></tr>")
                    .AppendLine($"<tr><td>貨運行</td><td>{ToHtmlString(dropship.freight)}</td></tr>")
                    .AppendLine("</table>")
                    .AppendLine("<p>若有問題請聯絡成倉管理者。</p>");


                emailWork = CreateWorkItem(recipientsList, ccList, fromAddress, subject, bodySb.ToString(), smtpServer, smtpPort, smtpUser, smtpPass);
                WebStoreHouse.Services.BackgroundEmailQueue.Instance.Enqueue(emailWork);
                await Task.Yield();
            }
            catch (Exception ex)
            {
                LogEmailException(ex, nameof(SendDropshipDeletedEmailAsync), smtpServer, smtpPort, smtpUser, fromAddress, recipientsList, ccList, subject);
                throw;
            }
        }

        // 將常用的 Html encode 與 null-safe Convert 包成小方法
        /// <summary>
        /// 將指定物件轉為 HTML 編碼字串，避免 XSS 攻擊與特殊字元問題。
        /// </summary>
        /// <param name="value">要轉換的物件，可為 null。</param>
        /// <returns>HTML 編碼後的字串，若 value 為 null 則回傳空字串。</returns>
        private static string ToHtmlString(object value)
        {
            // WHY: 統一 HTML 編碼，避免 XSS 與特殊字元問題
            return System.Web.HttpUtility.HtmlEncode(value?.ToString() ?? string.Empty);
        }

        #region 收件者清單
        /// <summary>
        /// 取得收件者清單（可改為從 DB 或設定檔讀取）
        /// </summary>
        /// <param name="dropship">直出資料物件</param>
        /// <returns>收件者 Email 清單</returns>
        public virtual IEnumerable<string> GetRecipientsForDropship(dynamic dropship)
        {
            // WHY: 可依需求改寫，預設空集合
            var list = new List<string>();
            list.Add("Vian.Shiu@eversun.com.tw"); // IE
            //list.Add("jesse.chiang@eversun.com.tw");
            return list.Distinct(StringComparer.OrdinalIgnoreCase);
        }

        /// <summary>
        /// 取得副本收件者（CC），預設為空集合。可由子類別覆寫以提供動態邏輯。
        /// </summary>
        /// <param name="dropship">直出資料物件。</param>
        /// <returns>CC 電子郵件地址集合。</returns>
        public virtual IEnumerable<string> GetCcForDropship(dynamic dropship)
        {
            // WHY: 預設 CC 為 MIS，可依需求擴充
            var list = new List<string>
            {
                "jesse.chiang@eversun.com.tw"
            };
            return list.Distinct(StringComparer.OrdinalIgnoreCase);
        }
        #endregion

        #region 紀錄 Email 發送時發生的例外狀況
        /// <summary>
        /// 紀錄 Email 發送時發生的例外狀況，將詳細資訊寫入 Log/email_errors.log 檔案。
        /// 此方法不會拋出例外，若紀錄失敗則靜默處理。
        /// </summary>
        /// <param name="ex">捕捉到的例外物件。</param>
        /// <param name="caller">呼叫此方法的函式名稱。</param>
        /// <param name="smtpServer">SMTP 伺服器位址。</param>
        /// <param name="smtpPort">SMTP 伺服器連接埠。</param>
        /// <param name="smtpUser">SMTP 使用者名稱。</param>
        /// <param name="fromAddress">寄件者電子郵件地址。</param>
        /// <param name="recipients">收件者電子郵件地址集合。</param>
        /// <param name="subject">郵件主旨。</param>
        private static void LogEmailException(Exception ex, string caller, string smtpServer, int smtpPort, string smtpUser, string fromAddress, IEnumerable<string> recipients, IEnumerable<string> cc, string subject)
        {
            try
            {
                var basePath = HttpContext.Current?.Server.MapPath("~/Log") ?? Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Log");
                Directory.CreateDirectory(basePath);
                var logPath = Path.Combine(basePath, $"email_errors_{DateTime.Now:yyyyMMdd}.log");
                var recipientsText = recipients != null ? string.Join(";", recipients) : "(none)";
                var ccText = cc != null ? string.Join(";", cc) : "(none)";

                var exText = ex.ToString();

                string locationInfo = null;
                try
                {
                    var st = new System.Diagnostics.StackTrace(ex, true);
                    var frame = st.GetFrames()?.FirstOrDefault();
                    if (frame != null)
                    {
                        var file = frame.GetFileName();
                        var line = frame.GetFileLineNumber();
                        if (!string.IsNullOrEmpty(file) && line > 0)
                            locationInfo = $" at {file}:{line}";
                    }
                }
                catch { }

                var log = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {caller}{locationInfo} - SMTP:{smtpServer}:{smtpPort} User:{smtpUser} From:{fromAddress} To:{recipientsText} Cc:{ccText} Subject:{subject}\nException:{exText}\n\n";
                File.AppendAllText(logPath, log);
            }
            catch
            {
                // 紀錄失敗不應影響主要例外流程
            }
        }

        #endregion
    }
}
