using System.Web;
using System.Web.Optimization;

namespace WebStoreHouse
{
    public class BundleConfig
    {
        // 如需統合的詳細資訊，請瀏覽 https://go.microsoft.com/fwlink/?LinkId=301862
        public static void RegisterBundles(BundleCollection bundles)
        {
            // 使用已經最小化的檔案，並改用 plain Bundle 避免內建 JsMinify 解析錯誤
            // 將 bundle 指向專案中實際存在的 jQuery 檔案，避免指向不存在的檔案導致 bootstrap 在沒有 jQuery 時被載入
            bundles.Add(new Bundle("~/bundles/jquery").Include(
                        "~/Scripts/jquery-3.4.1.min.js"));

            bundles.Add(new ScriptBundle("~/bundles/jqueryval").Include(
                        "~/Scripts/jquery.validate*"));

            // 使用開發版本的 Modernizr 進行開發並學習。然後，當您
            // 準備好可進行生產時，請使用 https://modernizr.com 的建置工具，只挑選您需要的測試。
            bundles.Add(new ScriptBundle("~/bundles/modernizr").Include(
                        "~/Scripts/modernizr-*"));

            bundles.Add(new Bundle("~/bundles/bootstrap").Include(
                      "~/Scripts/bootstrap.min.js"));

            bundles.Add(new StyleBundle("~/Content/css").Include(
                      "~/Content/bootstrap.css",
                      "~/Content/site.css"));

            // 臨時修正：停用最佳化（minification/concat），避免內建 Microsoft Ajax Minifier
            // 在處理較新或含有 ES6+ 語法的 JS 檔案（如新版 jQuery/Bootstrap）時，會在 JsMinify 解析器中發生 NullReferenceException。
            // 若要在生產環境使用 minify，建議安裝並改為使用 NUglify（或其他現代 minifier）來取代內建的 JsMinify。
            // 範例：安裝 NuGet 套件 "BuildBundlerMinifier" 或 "NUglify"，並將 transform 換成 NUglify 的 minifier。
            // 目前先預設關閉最佳化以立即避免例外。
            BundleTable.EnableOptimizations = false;
        }
    }
}
