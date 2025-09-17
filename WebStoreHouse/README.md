# 成品倉 WebStoreHouse

本文件為專案操作與維運說明的集中文件，目標是讓開發者、維運或新進人員快速理解系統架構、開發與部署流程，以及日常常見維運任務。

## 簡短系統概述

成品倉 WebStoreHouse 是一個以 ASP.NET MVC（.NET Framework）為基礎的 Web 應用，專案結構採傳統 MVC 分層設計：

- 控制器：`Controllers/`
- 模型：`Models/` 與 `ViewModels/`
- 視圖（UI）：`Views/`
- 靜態資源：`Content/`、`Scripts/`、`Images/`、`fonts/`
- 專案入口與設定：`Global.asax`、`Web.config`、`WebStoreHouse.sln`

資料存取方面專案中含有 EntityFramework（EF6）與 Dapper，因此可能同時使用 ORM 與輕量 SQL 查詢方式處理資料庫互動。若需調整資料庫連線設定，請檢查 `Web.config` 中的 `<connectionStrings>` 區塊。

> 假設：本說明假設專案為 .NET Framework (4.x) 的 ASP.NET MVC 應用，資料庫為 Microsoft SQL Server。如非此設定請告知，我會依實際情況更新說明。

## 目錄與重點檔案

- `WebStoreHouse.sln`：解決方案檔，開啟此檔即可在 Visual Studio 中載入所有專案檔案。
- `Web.config`：主要的應用程式設定（包含 connectionStrings、appSettings 等）。
- `packages/`：本地 NuGet 套件快取（若您使用 Visual Studio 直接還原套件，通常不需額外下載）。
- `SqlServerTypes/`：若專案使用 SQL Server 的地理空間或特定型別，請確認 Native assemblies 是否正確載入。

---

## 在開發環境執行（本機）

建議使用 Visual Studio 2019 或 2022（支援 .NET Framework 的版本）。基本步驟：

1. 使用 Visual Studio 開啟 `WebStoreHouse.sln`。
2. 還原 NuGet 套件：Visual Studio 會在開啟時自動還原；若要使用命令列：

    ```powershell
    # 在專案根目錄（含 WebStoreHouse.sln）執行（需先安裝 nuget.exe 或使用 Visual Studio 的還原）
    nuget restore .\WebStoreHouse.sln
    msbuild .\WebStoreHouse.sln /p:Configuration=Debug
    ```

3. 設定資料庫連線：編輯 `Web.config` 中的 `<connectionStrings>`，把開發環境的連線字串填入（例如本機 SQL Server 或 Dockerized SQL）。
4. 在 Visual Studio 中選擇 IIS Express 或本機 IIS，啟動 Debug（F5）或直接執行（Ctrl+F5）。

### 常見本機問題與解法

- 如果啟動失敗且顯示「無法連線資料庫」，請先確認 SQL Server 正在執行、連線字串無誤，以及應用程式帳號有存取權。可用 SQL Server Management Studio 測試連線。
- 權限問題（例如寫入 `App_Data`）：請確認專案所在資料夾對應的 IIS 使用者或執行帳戶擁有讀寫權限。

---

## 系統流程說明（高層）

1. 使用者透過瀏覽器發出 HTTP 請求。
2. 路由系統（RouteConfig）決定要呼叫哪個 `Controller` 與 `Action`。
3. Controller 負責處理請求、呼叫商業邏輯或資料層（Models / Service），並準備資料給 View 或回傳 JSON。
4. View（Razor）將資料渲染為 HTML，瀏覽器載入靜態資源（JavaScript、CSS、圖片）以完成頁面。

資料層可能會使用：

- Entity Framework 6（DbContext、Repository pattern）
- Dapper（直接 SQL 查詢以提高效能情境）

如果要追蹤某筆功能流程，建議：先看 `Controllers/` 的對應 Controller，再追到 `Services/` 或 `Models/`，最後檢視 Repository/DB 查詢。

---

## 部署到 IIS（生產環境）

基本部署流程：

1. 在伺服器上建立 IIS Site（或 Application）並設定 Application Pool 使用 .NET CLR v4.0。
2. 發佈專案：在 Visual Studio 使用「發佈」功能或手動將編譯後的內容（`bin/`、`Views/`、`Content/`、`Scripts/`、`Web.config` 等）部署到網站根目錄。
3. 更新 `Web.config` 的生產環境 `connectionStrings` 與必要的 `appSettings`。
4. 確保伺服器有安裝相對應的 Visual C++ Redistributable（若 native assemblies 需要），以及 SQL Server 存取權限與網路連線。

建議的 IIS 設定：

- Application Pool Identity：使用專用服務帳號（非預設 LocalSystem）以便更細緻控制資料庫與檔案權限。
- 針對靜態檔案啟用長期快取（Cache-Control）以提升效能。

---

## 資料庫與資料遷移

- 專案使用 EF6；若使用 Code First Migration，請在開發環境中使用 `Update-Database`（Package Manager Console）來同步資料庫結構。若專案未採用 migration，請依備註走 SQL 腳本手動建立表格。
- 資料庫連線字串位置：`Web.config` 的 `<connectionStrings>`。

---

## 日誌、錯誤追蹤與除錯建議

- 若專案沒有整合第三方日誌（如 Serilog / NLog），可暫時在 Controller 中捕捉例外並觀察 `Event Viewer` 或在 `bin/` 上設定更詳細的錯誤頁面（僅在測試環境）。
- 建議導入集中式日誌與錯誤回報（例如：Application Insights、ELK、Sentry），以利生產環境追蹤。

---

## 常見維運任務（Checklist）

- 更新 NuGet 套件：在 Visual Studio 的 NuGet 管理工具中檢查可用更新，先在開發分支測試再上線。
- 建置並部署：使用 CI/CD（例如 GitLab CI、Azure DevOps）把 build 與 deploy 自動化，避免手動錯誤。
- 備份資料庫：建立定期備份策略並驗證還原流程。

---

## Troubleshooting（快速排查表）

- 500 系錯誤：檢查 `Web.config` 的 customErrors 與伺服器事件日誌，啟用詳細錯誤以便排查（僅在非生產環境）。
- 靜態檔案 404：確認 `Content/`、`Scripts/` 路徑是否正確以及 bundling 設定（若有使用 BundleConfig）。
- 一致性錯誤（Concurrency）：檢查資料庫鎖、交易邏輯以及 EF 的 SaveChanges 行為。

---

## 開發者建議與最佳實務

- 在修改資料庫結構前，先在 staging 或本機執行 migration 與回歸測試。
- 建議把敏感設定（如 connection strings）在部署時以環境變數或部署時注入的方式管理，而非直接在版本庫中保留明文憑證。
- 寫單元測試與整合測試（至少針對商業邏輯層與資料層）以降低上線風險。

---

## 假設與未來工作建議

- 假設：專案為 ASP.NET MVC (.NET Framework) + SQL Server，若您使用不同版本或改成 .NET Core/.NET 6+，我可以協助移轉計畫。
- 建議未來可加入：CI/CD 自動化發佈、集中式日誌、單元測試套件、以及部署腳本（PowerShell 或 Azure/Octopus 等）。

---

## 聯絡與貢獻

若您要我協助：

- 新增「一鍵部署」或 CI/CD 範本
- 撰寫資料庫初始化腳本或 migration
- 整合日誌/監控

請回覆您要我優先協助的項目，我可以直接在專案中建立相關腳本與設定範本。

---

**DISCLAIMER**: 這份 README 旨在幫助開發與維運；如需針對專案細節（例如確切的 .NET Framework 版本或資料庫腳本）進行補充，請提供 `Web.config` 中的連線字串樣本或專案的 `App_Start` 內容，我會把文件更新得更精準。
