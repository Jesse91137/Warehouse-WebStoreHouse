using System;
using System.Collections.Generic;
using System.Linq;
using System.Web.Mvc;
using System.Web.Optimization;
using System.Web.Routing;
using System.Reflection;
using Autofac;

namespace WebStoreHouse
{
    public class MvcApplication : System.Web.HttpApplication
    {
        /// <summary>
        /// ASP.NET MVC 應用程式啟動時執行的事件處理函式。
        /// 負責註冊 MVC 區域、全域篩選器、路由、資源綁定，
        /// 並自動掃描並註冊 Controller、Service、Repository 至 Autofac DI 容器。
        /// 若有安裝 Autofac.Integration.Mvc 則使用官方整合，否則使用自訂的 AutofacResolver。
        /// 最後註冊每週日午夜自動清理 Log 檔案的排程。
        /// </summary>
        protected void Application_Start()
        {
            AreaRegistration.RegisterAllAreas();
            FilterConfig.RegisterGlobalFilters(GlobalFilters.Filters);
            RouteConfig.RegisterRoutes(RouteTable.Routes);
            BundleConfig.RegisterBundles(BundleTable.Bundles);

            #region 自動掃描並註冊 新增 20251008 by Jesse
            // Autofac 容器設定：自動掃描並註冊 controller、service、repository
            var builder = new ContainerBuilder();

            // 手動註冊所有 MVC controllers（避免依賴 Autofac.Integration.Mvc）
            var thisAssembly = Assembly.GetExecutingAssembly();
            var controllerTypes = thisAssembly.GetTypes()
                .Where(t => typeof(Controller).IsAssignableFrom(t) && !t.IsAbstract)
                .ToArray();
            if (controllerTypes.Length > 0)
            {
                builder.RegisterTypes(controllerTypes).AsSelf().InstancePerDependency();
            }

            // 只掃描應用程式相關的 assemblies，避免註冊系統或第三方 framework types（例如 WorkflowDebuggerService）
            // 只包含非動態、具有 Location，且位於應用程式目錄的 assemblies（或目前 assembly）
            var assemblies = AppDomain.CurrentDomain.GetAssemblies()
             .Where(a => !a.IsDynamic && !string.IsNullOrEmpty(a.Location) &&
                   (a == thisAssembly || a.Location.StartsWith(AppDomain.CurrentDomain.BaseDirectory, StringComparison.OrdinalIgnoreCase)))
             .ToArray();

            // 範例：掃描並將名稱以 Service 或 Repository 結尾，且為可實作介面的具名類別註冊為其介面
            builder.RegisterAssemblyTypes(assemblies)
                .Where(t => (t.Name.EndsWith("Service") || t.Name.EndsWith("Repository"))
                      && t.IsClass && !t.IsAbstract && t.IsPublic
                      && t.GetInterfaces().Length > 0)
                .AsImplementedInterfaces()
                .InstancePerDependency(); // 使用 InstancePerDependency 為預設範圍

            // 如果有安裝 Autofac MVC Integration，可以改用官方擴充以支援 InstancePerRequest() 與 FilterProvider。
            // 如果尚未安裝，程式會使用下方自訂的 AutofacResolver 作為 fallback。
            var container = builder.Build();

            // 嘗試使用官方整合型別（若專案已引用 Autofac.Integration.Mvc）
            try
            {
                // 使用 reflection 檢查是否有 Autofac.Integration.Mvc 的類型
                var integrationType = Type.GetType("Autofac.Integration.Mvc.AutofacDependencyResolver, Autofac.Integration.Mvc");
                var registerControllersMethod = Type.GetType("Autofac.Integration.Mvc.RegistrationExtensions, Autofac.Integration.Mvc");
                if (integrationType != null)
                {
                    // 若存在官方整合則使用官方 DependencyResolver
                    var resolver = Activator.CreateInstance(integrationType, container) as IDependencyResolver;
                    if (resolver != null)
                    {
                        DependencyResolver.SetResolver(resolver);
                        return;
                    }
                }
            }
            catch
            {
                // 忽略，使用 fallback
            }

            // fallback 使用自訂 adapter
            DependencyResolver.SetResolver(new AutofacResolver(container));

            #endregion

            // 註冊每週日午夜的 Log 清理工作（刪除 Log\email_errors.log）
            try
            {
                RegisterWeeklyLogCleanup();
            }
            catch
            {
                // 若註冊失敗，不影響 App 啟動
            }
        }

        #region 定時清理 Log 新增 20251008 by Jesse
        /// <summary>
        /// 每週日午夜執行 Log 清理工作的計時器實例。
        /// 用於定時檢查並刪除 Log\email_errors.log 檔案。
        /// </summary>
        private System.Threading.Timer _weeklyCleanupTimer;

        /// <summary>
        /// 註冊一個 Timer，每分鐘檢查一次，如果目前為週日 00:00 就執行清理
        /// </summary>
        private void RegisterWeeklyLogCleanup()
        {
            // 每分鐘檢查一次是否到達週日 00:00
            var dueTime = TimeSpan.Zero; // 立即啟動
            var period = TimeSpan.FromMinutes(1);
            _weeklyCleanupTimer = new System.Threading.Timer(state =>
            {
                try
                {
                    var now = DateTime.Now;
                    // 檢查是否為週日且時間在 00:00 到 00:01 內（允許一分鐘容忍）
                    if (now.DayOfWeek == DayOfWeek.Sunday && now.Hour == 0 && now.Minute == 0)
                    {
                        var logDir = System.Web.Hosting.HostingEnvironment.MapPath("~/Log");
                        if (!string.IsNullOrEmpty(logDir))
                        {
                            var logPath = System.IO.Path.Combine(logDir, "email_errors.log");
                            if (System.IO.File.Exists(logPath))
                            {
                                try
                                {
                                    System.IO.File.Delete(logPath);
                                }
                                catch { }
                            }
                        }
                    }
                }
                catch
                {
                    // 忽略所有清理錯誤
                }
            }, null, dueTime, period);
        }
        #endregion

        /// <summary>
        /// AutofacResolver 是一個簡易的 Autofac 容器與 ASP.NET MVC 的 <see cref="IDependencyResolver"/> 介面之間的適配器。
        /// 允許在未安裝 Autofac.Integration.Mvc 的情況下，仍可使用 Autofac 作為 MVC 的 DI 容器。
        /// </summary>
        private class AutofacResolver : IDependencyResolver, IDisposable
        {
            /// <summary>
            /// Autofac DI 容器實例。
            /// </summary>
            private readonly IContainer container;

            /// <summary>
            /// 建構函式，初始化 AutofacResolver 並指定 DI 容器。
            /// </summary>
            /// <param name="container">要使用的 Autofac DI 容器。</param>
            /// <exception cref="ArgumentNullException">若 <paramref name="container"/> 為 null 則拋出例外。</exception>
            public AutofacResolver(IContainer container)
            {
                this.container = container ?? throw new ArgumentNullException(nameof(container));
            }

            /// <summary>
            /// 取得指定型別的服務實例。
            /// </summary>
            /// <param name="serviceType">要解析的服務型別。</param>
            /// <returns>解析成功則回傳服務實例，否則回傳 null。</returns>
            public object GetService(Type serviceType)
            {
                if (serviceType == null) return null;
                try
                {
                    if (container.IsRegistered(serviceType))
                        return container.Resolve(serviceType);
                }
                catch
                {
                    // resolve 失敗回傳 null，讓 MVC 處理 fallback
                }
                return null;
            }

            /// <summary>
            /// 取得指定型別的所有服務實例集合。
            /// </summary>
            /// <param name="serviceType">要解析的服務型別。</param>
            /// <returns>解析成功則回傳服務集合，否則回傳空集合。</returns>
            public IEnumerable<object> GetServices(Type serviceType)
            {
                if (serviceType == null) return Enumerable.Empty<object>();
                var enumerableService = typeof(IEnumerable<>).MakeGenericType(serviceType);
                try
                {
                    if (container.IsRegistered(enumerableService))
                    {
                        var resolved = (IEnumerable<object>)container.Resolve(enumerableService);
                        return resolved ?? Enumerable.Empty<object>();
                    }
                }
                catch
                {
                    // ignore
                }
                return Enumerable.Empty<object>();
            }

            /// <summary>
            /// 釋放 Autofac DI 容器資源。
            /// </summary>
            public void Dispose()
            {
                try { container.Dispose(); } catch { }
            }
        }

        // 新增 20251008 by Jesse
        /// <summary>
        /// 當每個 HTTP 請求結束時觸發的事件處理函式。
        /// 若回應狀態碼為 404，且請求 URL 不包含 "style" 或 "images"，
        /// 則清除回應並導向至 /Home/Login 頁面。
        /// </summary>
        protected void Application_EndRequest()
        {
            if (Context.Response.StatusCode == 404)
            {
                if ((!Request.RawUrl.Contains("style")) && (!Request.RawUrl.Contains("images")))
                {
                    Response.Clear();
                    if (Response.StatusCode == 404)
                    {
                        Response.Redirect("/Home/Login");
                    }
                }
            }
        }

        /// <summary>
        /// 當 Session 開始時觸發的事件處理函式。
        /// 設定 Session 的存活時間為 480 分鐘。
        /// 新增 20250807 by Jesse
        /// </summary>
        /// <param name="sender">事件來源物件。</param>
        /// <param name="e">事件相關的參數。</param>
        protected void Session_Start(object sender, EventArgs e)
        {
            Session.Timeout = 480; // 設定 Session 保持時間為 480 分鐘
        }

    }
}
