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

---

## 系統操作詳解（給開發者與維運）

以下章節提供更具體、可執行的檢查清單與操作步驟，讓開發者或維運人員能在本機、測試或生產環境中快速啟動、偵錯與排除問題。

### 先決條件（Prerequisites）

- Windows 10/Server（支援 IIS）
- Visual Studio 2019 或 2022（支援 .NET Framework）
- SQL Server（或相容 DB）可供連線
- nuget.exe（若需從命令列還原）
- 必要的 native assemblies（例如 `SqlServerTypes` 對應的 x86/x64 DLL）已放置於專案或部署檔案中

### 啟動與本機運行（詳細步驟）

1. 取得程式碼並還原套件

     - 在專案根目錄（含 `WebStoreHouse.sln`）執行還原：

         ```powershell
         nuget restore .\WebStoreHouse.sln
         msbuild .\WebStoreHouse.sln /p:Configuration=Debug
         ```

     - 或在 Visual Studio 開啟 `WebStoreHouse.sln`，IDE 會自動還原 NuGet 套件。

2. 設定資料庫連線

     - 編輯 `Web.config` 中的 `<connectionStrings>`。建議在 local 開發使用本機 SQL Server（或 Docker container）測試。
     - 若使用開發/測試/生產多個環境，請採用 Web.config transform 或在部署時注入連線字串（避免把敏感資訊放入版本庫）。

3. 啟動應用程式

     - 建議以 Visual Studio 的 IIS Express（F5）或使用本機 IIS（需要先設定 Site 與 Application Pool）：

         - 如果使用本機 IIS：
             1. 在 IIS 管理員建立一個 Site 或 Application，指向專案發佈輸出資料夾。
             2. Application Pool 使用 .NET CLR v4.0 並設定適當的 Identity（建議專用服務帳號）。

     - 若需發佈（Publish）：在 Visual Studio 使用「發佈」功能，或手動把編譯後的 `bin/`、`Views/`、`Content/`、`Scripts/`、`Web.config` 複製到網站目錄。

### 快速檢查清單（當網站無法啟動）

- 檢查 `Web.config` 的 `connectionStrings` 是否正確。
- 確認 SQL Server 可連線（可用 SSMS 測試）。
- 檢查 IIS Application Pool 是否啟動，並確認使用的 .NET CLR 版本。
- 檢查網站檔案權限（例如 `App_Data` 是否有寫入權限）。
- 查看 Windows Event Viewer 與 IIS 日誌（或開啟 Failed Request Tracing）取得詳細錯誤碼與堆疊。

### 請求處理流程（Request flow — 詳細解說）

1. 使用者發出 HTTP 請求到 URL。
2. IIS 接受請求並依設定將請求轉給 ASP.NET MVC 管線。
3. RouteConfig（或 Attribute Routing）解析路由並決定要呼叫哪個 Controller/Action。
4. Controller 解析輸入（Model Binding）、驗證授權，並呼叫 Service 層或商業邏輯。
5. Service 層或 Manager 層負責處理業務邏輯，並呼叫資料存取層（Repository）進行資料查詢或寫入。
6. 資料層可能使用 Entity Framework（DbContext）或 Dapper（手寫 SQL）進行 DB 交互。
7. Controller 根據需求回傳 View（Razor）或 JSON（API），視圖負責把資料渲染成 HTML 並回傳給瀏覽器。

在除錯時，請沿著上述流程從 Controller 開始往下追蹤（先看路由、再看 Controller、Service、Repository、最後是 SQL）。

### 資料庫與遷移（EF6）

- 如果專案採用 EF Code First 與 Migration，使用 Package Manager Console（PMC）：

    - 啟用 migration（只需做一次）：

        ```powershell
        Enable-Migrations
        Add-Migration InitialCreate
        Update-Database
        ```

    - 若僅在生產環境更新 schema，請先在 staging 測試 `Update-Database`，並將產生的 SQL script 備份。

- 備份與還原（以 SQL Server 為例）：

    - 備份（T-SQL）：

        ```sql
        BACKUP DATABASE [YourDatabase] TO DISK = N'C:\Backups\YourDatabase_full.bak' WITH INIT, FORMAT;
        ```

    - 還原（T-SQL）：

        ```sql
        RESTORE DATABASE [YourDatabase] FROM DISK = N'C:\Backups\YourDatabase_full.bak' WITH REPLACE;
        ```

    - 建議使用 SQL Server Agent 排程週期性備份，並驗證還原流程。

### SqlServerTypes 與 Native Assemblies

若專案使用 SQL Server 之地理空間型別或其他 native assemblies，需把對應的 x86/x64 DLL 放入 `bin`，並在啟動時載入（範例放在 Global.asax）：

```csharp
// ...existing code...
// 在 Application_Start 或適合的位置呼叫：
SqlServerTypes.Utilities.LoadNativeAssemblies(Server.MapPath("~"));
// ...existing code...
```

### 日誌、錯誤追蹤與診斷（Operational diagnostics）

- 檢查點：
    - IIS 日誌位於 `%SystemDrive%\inetpub\logs\LogFiles`。
    - Windows Event Viewer 可查看應用程式層級錯誤。
    - 若專案有自訂日誌（例如 log4net、NLog、Serilog），請確認 `App_Data` 或日誌目錄權限並查看最近檔案。

- 建議：
    - 在 Global.asax 的 `Application_Error` 中把未處理例外記錄到檔案或外部系統。
    - 生產環境建議導入集中式日誌（Application Insights、ELK、Seq、Sentry 等）。

### 常見故障與處理步驟（Runbook）

1. 500 系列錯誤（網站回傳 500）

    - 步驟：
        1. 開啟 `Web.config` 的 `customErrors` 設定為 `Off`（僅在測試環境），或查看 Event Viewer 與 IIS Failed Request Tracing。
        2. 檢查 `Global.asax` 的 `Application_Error` 是否有記錄例外。
        3. 檢查內部 exception 的 StackTrace，定位 Controller 或 Service 層的例外來源。

2. 無法連線資料庫（連線逾時或拒絕）

    - 步驟：
        1. 用 SSMS 測試連線字串是否能連線到 DB。
        2. 檢查防火牆、SQL Server 是否允許遠端連線，或是否有登入權限問題。
        3. 在 `Web.config` 中確認帳號與密碼是否正確（若在生產環境請用安全的儲存方式）。

3. 靜態檔 404（CSS / JS / 圖片）

    - 步驟：
        1. 檢查資源路徑是否正確，確認是否在 `Content/` 或 `Scripts/` 中存在該檔案。
        2. 如果使用 BundleConfig，確認 bundle 設定與 `BundleTable.EnableOptimizations` 在開發/生產模式的差異。

4. 權限問題（檔案寫入失敗）

    - 步驟：
        1. 檢查網站目錄與 `App_Data` 的 NTFS 權限，確保 Application Pool Identity 有寫入權限。
        2. 若使用 Windows Service 帳號，確認該帳號擁有必要的權限。

### 背景工作與排程任務

- 若專案使用 Background Job（例如 Hangfire），請確認資料庫連線字串與 Background Job Server 是否在可用的主機上執行。
- 如果使用 Windows Task Scheduler 或類似工具排程，請確認 Task 的執行帳號與工作目錄。

### 建議的觀察指標與監控（要監控什麼）

- 可用性（HTTP 2xx / 4xx / 5xx 比例）
- 平均回應時間、P95/P99 延遲
- 錯誤率（例外量）
- 資料庫長查詢與死鎖
- 磁碟與記憶體使用率、GC 次數（若可取得）

建議工具：Prometheus + Grafana、Application Insights、ELK、Datadog。

### 維運常用命令與操作清單

- 重新啟動 App Pool（IIS）：在伺服器上使用 IIS 管理員或 PowerShell 指令重啟 App Pool。

    ```powershell
    Import-Module WebAdministration; Restart-WebAppPool -Name "YourAppPoolName"
    ```

- 查看最近 N 行 IIS 日誌（PowerShell 範例）：

    ```powershell
    Get-Content -Path 'C:\inetpub\logs\LogFiles\W3SVC1\u_ex*.log' -Tail 200
    ```

### 建議的改進與未來項目

- 新增健康檢查 endpoint（例如 `/health` 或 `/api/health`），回傳簡單的 DB 連線檢查與磁碟空間狀態。
- 導入集中式日誌與 APM（Application Performance Monitoring）。
- 建立 CI/CD 流程（例如 Azure DevOps / GitHub Actions）執行自動化建置、單元測試與發佈。

---

如果你要我代為建立：

- 一份測試用的 `Update-Database` SQL script 及 migration 範例；或
- 一個簡單的 `health` API endpoint 與部署指引；或
- 一個基礎的 CI/CD pipeline 範本（例如 Azure DevOps 或 GitHub Actions）

請告訴我你要優先哪一項，我會幫你建立對應的檔案與示範。
