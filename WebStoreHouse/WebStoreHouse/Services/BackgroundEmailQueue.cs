using System;
using System.Collections.Concurrent;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Threading;
using System.Threading.Tasks;
using System.Web;
using WebStoreHouse.Models;

namespace WebStoreHouse.Services
{
    /// <summary>
    /// 背景郵件佇列，單例、非阻塞地處理傳送請求
    /// 新增 20251008 by Jesse
    /// </summary>
    public sealed class BackgroundEmailQueue : IDisposable
    {
        /// <summary>
        /// 儲存待處理的郵件工作項目的佇列，使用 ConcurrentQueue 以確保執行緒安全。
        /// </summary>
        private readonly ConcurrentQueue<EmailWorkItem> _queue = new ConcurrentQueue<EmailWorkItem>();

        /// <summary>
        /// 用於通知背景工作有新郵件進入佇列的信號量。
        /// </summary>
        private readonly SemaphoreSlim _signal = new SemaphoreSlim(0);

        /// <summary>
        /// 控制背景工作取消執行的取消權杖來源。
        /// </summary>
        private readonly CancellationTokenSource _cts = new CancellationTokenSource();

        /// <summary>
        /// 背景處理郵件佇列的工作執行緒。
        /// </summary>
        private readonly Task _worker;

        /// <summary>
        /// 提供 BackgroundEmailQueue 的單例存取實例。
        /// </summary>
        public static BackgroundEmailQueue Instance { get; } = new BackgroundEmailQueue();

        /// <summary>
        /// 建構函式，初始化並啟動背景工作執行緒。
        /// </summary>
        private BackgroundEmailQueue()
        {
            // 啟動背景工作
            _worker = Task.Run(ProcessQueueAsync);
        }

        /// <summary>
        /// 將新的郵件工作項目加入佇列，並通知背景工作有新項目。
        /// </summary>
        /// <param name="item">要加入佇列的郵件工作項目。</param>
        public void Enqueue(EmailWorkItem item)
        {
            // 檢查郵件工作項目是否為 null
            if (item == null) return;
            // 將郵件工作項目加入佇列
            _queue.Enqueue(item);
            // 釋放信號量，通知背景工作有新項目
            _signal.Release();
        }

        /// <summary>
        /// 處理背景郵件佇列的非同步方法。
        /// 會持續監控佇列，並將郵件逐一傳送，失敗時記錄錯誤日誌。
        /// </summary>
        /// <returns>非同步 Task。</returns>
        private async Task ProcessQueueAsync()
        {
            // 持續執行，直到取消請求
            while (!_cts.IsCancellationRequested)
            {
                try
                {
                    // 等待有新郵件工作進入佇列
                    await _signal.WaitAsync(_cts.Token).ConfigureAwait(false);

                    // 只要佇列有工作就持續處理
                    while (_queue.TryDequeue(out var item))
                    {
                        try
                        {
                            // 非同步傳送郵件
                            await SendItemAsync(item).ConfigureAwait(false);
                        }
                        catch (Exception ex)
                        {
                            // 傳送失敗時，記錄錯誤日誌
                            try
                            {
                                // 取得日誌目錄路徑
                                var basePath = HttpContext.Current?.Server.MapPath("~/Log") ?? Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Log");
                                // 確保日誌目錄存在
                                Directory.CreateDirectory(basePath);
                                // 設定日誌檔案路徑
                                var logPath = Path.Combine(basePath, $"email_errors_{DateTime.Now:yyyyMMdd}.log");
                                // 取得收件者文字
                                var recipientsText = item?.To != null ? string.Join(";", item.To) : "(none)";
                                var ccText = item?.Cc != null ? string.Join(";", item.Cc) : "(none)";
                                // 取得例外文字
                                var exText = ex.ToString();
                                // 組合日誌內容
                                var log = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] BackgroundEmailQueue - SMTP:{item?.SmtpServer}:{item?.SmtpPort} User:{item?.SmtpUser} From:{item?.From} To:{recipientsText} Cc:{ccText} Subject:{item?.Subject}\nException:{exText}\n\n";
                                // 寫入日誌檔案
                                File.AppendAllText(logPath, log);
                            }
                            catch { /* 忽略日誌寫入失敗 */ }
                        }
                    }
                }
                // 當取消操作時捕捉例外
                catch (OperationCanceledException) when (_cts.IsCancellationRequested)
                {
                    // 允許取消，不做任何處理
                }
                // 捕捉其他例外，忽略並繼續執行
                catch
                {
                    // 忽略單次迴圈錯誤，繼續處理後續項目
                }
            }
        }

        /// <summary>
        /// 非同步傳送單一郵件工作項目。
        /// </summary>
        /// <param name="item">要傳送的郵件工作項目。</param>
        /// <returns>非同步 Task。</returns>
        private async Task SendItemAsync(EmailWorkItem item)
        {
            // 檢查郵件工作項目是否為 null
            if (item == null) return;

            // 建立 MailMessage 物件
            using (var msg = new MailMessage())
            {
                // 設定寄件者，若未指定則使用預設 no-reply
                msg.From = new MailAddress(item.From ?? "no-reply@eversun.com.tw");
                // 逐一加入收件者
                foreach (var to in item.To ?? Enumerable.Empty<string>())
                {
                    // 檢查收件者是否為空白
                    if (string.IsNullOrWhiteSpace(to)) continue;
                    try
                    {
                        // 加入收件者郵件地址
                        msg.To.Add(new MailAddress(to.Trim()));
                    }
                    catch
                    {
                        // 忽略格式錯誤的收件者
                    }
                }
                // 加入副本收件者 (CC)
                foreach (var cc in item.Cc ?? Enumerable.Empty<string>())
                {
                    if (string.IsNullOrWhiteSpace(cc)) continue;
                    try
                    {
                        msg.CC.Add(new MailAddress(cc.Trim()));
                    }
                    catch
                    {
                        // 忽略格式錯誤的 CC
                    }
                }
                // 若無有效收件者則不傳送
                if (!msg.To.Any() && !msg.CC.Any()) return;

                // 設定郵件主旨
                msg.Subject = item.Subject;
                // 設定郵件內容
                msg.Body = item.Body;
                // 設定是否為 HTML 格式
                msg.IsBodyHtml = item.IsBodyHtml;

                // 建立 SmtpClient 物件
                using (var client = new SmtpClient(item.SmtpServer, item.SmtpPort))
                {
                    // 若有指定帳號則設定憑證
                    if (!string.IsNullOrWhiteSpace(item.SmtpUser))
                        client.Credentials = new System.Net.NetworkCredential(item.SmtpUser, item.SmtpPass);
                    // 設定傳送方式為網路
                    client.DeliveryMethod = SmtpDeliveryMethod.Network;
                    // 設定逾時時間
                    client.Timeout = item.TimeoutMs > 0 ? item.TimeoutMs : 15000;
                    // 非同步傳送郵件
                    await client.SendMailAsync(msg).ConfigureAwait(false);
                }
            }
        }

        /// <summary>
        /// 釋放 BackgroundEmailQueue 所持有的資源。
        /// </summary>
        public void Dispose()
        {
            _cts.Cancel(); // 取消背景工作佇列的執行
            try
            {
                _worker?.Wait(3000); // 等待背景工作最多 3 秒結束
            }
            catch
            {
                // 忽略等待時發生的例外
            }
            _cts.Dispose(); // 釋放 CancellationTokenSource 資源
            _signal.Dispose(); // 釋放 SemaphoreSlim 資源
        }
    }
}
