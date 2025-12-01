using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using PagedList;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Entity;
using System.Data.SqlClient;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Web.Mvc;
using WebStoreHouse.Models;
using WebStoreHouse.Services;
using WebStoreHouse.ViewModels;

namespace WebStoreHouse.Controllers
{
    /// <summary>
    /// 倉庫庫存
    /// </summary>
    public class FunctionController : Controller
    {
        #region Fields
        /// <summary>
        /// 資料庫存取物件，提供 Entity Framework 資料操作功能。
        /// </summary>
        E_StoreHouseEntities db = new E_StoreHouseEntities();

        /// <summary>
        /// Email 通知服務介面，提供寄送 dropship 編輯/刪除通知的功能。
        /// </summary>
        private readonly IEmailNotificationService _emailService;

        /// <summary>
        /// 頁面分頁預設每頁筆數（倉庫庫存查詢）。
        /// </summary>
        int pagesize = 50;

        /// <summary>
        /// 入庫資料查詢每頁筆數。
        /// </summary>
        int entryPageSize = 30; // 入庫資料查詢每頁 30 筆
        #endregion

        #region Constructors
        /// <summary>
        /// 建構 FunctionController，注入 Email 通知服務。
        /// </summary>
        /// <param name="emailService">Email 通知服務介面。</param>
        public FunctionController(IEmailNotificationService emailService)
        {
            _emailService = emailService ?? throw new ArgumentNullException(nameof(emailService));
        }
        #endregion

        #region Action Methods
        // GET: Function
        /// <summary>
        /// Function
        /// </summary>
        /// <returns></returns>
        public ActionResult Index()
        {
            return View();
        }

        #region 倉庫庫存
        #region 查詢倉庫庫存
        //view
        /// <summary>
        /// 查詢倉庫庫存
        /// </summary>
        /// <param name="wono">工單號碼</param>
        /// <param name="order_cust">客戶業單</param>
        /// <param name="engsr">機種名稱</param>
        /// <param name="inputdate_start">入庫日期起</param>
        /// <param name="inputdate_end">入庫日期迄</param>
        /// <param name="page"></param>
        /// <returns></returns>
        public ActionResult StoreHouseStock(string wono, string order_cust, string engsr, string inputdate_start, string inputdate_end, int page = 1)
        {
            int currentPage = page < 1 ? 1 : page;
            //取得會員帳號指定fUserId
            var authResult = Login_Authentication();
            if (authResult != null)
            {
                return authResult;
            }
            List<StoreHouseStock> stores = new List<StoreHouseStock>();

            List<SqlParameter> parmL = new List<SqlParameter>();

            //最後將原始(serialno)位置庫存設定為0 設定del_flag = 'D', 設定為D不備搜尋出來, del備註自動填寫轉庫位

            // 20250821 因應歷史資料的入庫日期的呈現要完整日期，原本寫入入庫日期只存日期，改為存完整日期+時間，入庫日期修改顯示日期 By Jesse
            string strsql = @"select serialno,sno,nowono,wono,cust_wono,eng_sr,order_count,quantity,box_quantity,kf10,kq30,
                                sap_in,position,acc_in,outed,notout,borrow,due_date,mark,CONVERT(varchar(10), inputdate, 120) AS inputdate,package,output_local,Igroup 
                                from E_StoreHouseStock where quantity >0 and (del_flag is null or del_flag <> 'D') and (igroup ='' or Igroup is null)  ";
            if (!string.IsNullOrWhiteSpace(wono))
            {
                strsql += " and wono = @wono ";
                parmL.Add(new SqlParameter("wono", wono));
            }
            if (!string.IsNullOrWhiteSpace(order_cust))
            {
                strsql += " and cust_wono = @order_cust ";
                parmL.Add(new SqlParameter("order_cust", order_cust));
            }
            if (!string.IsNullOrWhiteSpace(engsr))
            {
                strsql += " and eng_sr = @engsr ";
                parmL.Add(new SqlParameter("engsr", engsr));
            }
            // 20250821 因應歷史資料的入庫日期的呈現要完整日期，原本寫入入庫日期只存日期，改為存完整日期+時間，入庫日期修改顯示日期 By Jesse
            if (!string.IsNullOrWhiteSpace(inputdate_start))
            {
                strsql += " and CONVERT(varchar(10), inputdate, 120) >= @inputdate_start ";
                parmL.Add(new SqlParameter("inputdate_start", inputdate_start));
            }
            if (!string.IsNullOrWhiteSpace(inputdate_end))
            {
                strsql += " and CONVERT(varchar(10), inputdate, 120) <= @inputdate_end ";
                parmL.Add(new SqlParameter("inputdate_end", inputdate_end));
            }
            // 20250821 因應歷史資料的入庫日期的呈現要完整日期，原本寫入入庫日期只存日期，改為存完整日期+時間，入庫日期修改顯示日期 By Jesse
            strsql += @" union
                        select serialno,sno,nowono,wono,cust_wono,eng_sr,order_count,quantity,box_quantity,kf10,kq30,
                        sap_in,position,acc_in,outed,notout,borrow,due_date,mark,CONVERT(varchar(10), inputdate, 120) AS inputdate,package,output_local,Igroup 
                        from (Select *,ROW_NUMBER() Over (Partition By Igroup Order By inputdate Desc) As Sort 
                        From E_StoreHouseStock where igroup <>'' and igroup is not null and (del_flag is null or del_flag <> 'D') 
                        and quantity >0) s where sort=1 ";
            if (!string.IsNullOrWhiteSpace(wono))
            {
                strsql += " and wono = @wono2 ";
                parmL.Add(new SqlParameter("wono2", wono));
            }
            if (!string.IsNullOrWhiteSpace(order_cust))
            {
                strsql += " and cust_wono = @order_cust2 ";
                parmL.Add(new SqlParameter("order_cust2", order_cust));
            }
            if (!string.IsNullOrWhiteSpace(engsr))
            {
                strsql += " and eng_sr = @engsr2 ";
                parmL.Add(new SqlParameter("engsr2", engsr));
            }
            // 20250821 因應歷史資料的入庫日期的呈現要完整日期，原本寫入入庫日期只存日期，改為存完整日期+時間，入庫日期修改顯示日期 By Jesse
            if (!string.IsNullOrWhiteSpace(inputdate_start))
            {
                strsql += " and CONVERT(varchar(10), inputdate, 120) >= @inputdate_start2 ";
                parmL.Add(new SqlParameter("inputdate_start2", inputdate_start));
            }
            if (!string.IsNullOrWhiteSpace(inputdate_end))
            {
                strsql += " and CONVERT(varchar(10), inputdate, 120) <= @inputdate_end2 ";
                parmL.Add(new SqlParameter("inputdate_end2", inputdate_end));
            }

            // 修正：SQL 查詢最後加上排序，確保 union 後資料正確排序
            strsql += " order by eng_sr ";

            SqlDataReader dr = dbMethod.ExecuteReaderPmsList(strsql, CommandType.Text, parmL);
            while (dr.Read())
            {
                StoreHouseStock store = new StoreHouseStock();
                store.serialno = (int)dr["serialno"];
                store.nowono = dr["nowono"].ToString();
                store.sno = (int)dr["sno"];
                store.wono = dr["wono"].ToString();
                store.cust_wono = dr["cust_wono"].ToString();
                store.eng_sr = dr["eng_sr"].ToString();
                store.order_count = (int)dr["order_count"];
                store.quantity = (int)dr["quantity"];
                store.box_quantity = (int)dr["box_quantity"];
                store.kf10 = (int)dr["kf10"];
                store.kq30 = (int)dr["kq30"];
                store.sap_in = dr["sap_in"].ToString();
                store.position = dr["position"].ToString();
                store.acc_in = (int)dr["acc_in"];
                store.outed = (int)dr["outed"];
                store.notout = (int)dr["notout"];
                store.borrow = (int)dr["borrow"];
                store.due_date = (DateTime)dr["due_date"];
                store.mark = dr["mark"].ToString();
                store.inputdate = dr["inputdate"].ToString();
                store.package = dr["package"].ToString();
                store.output_local = dr["output_local"].ToString();
                store.Igroup = dr["Igroup"].ToString();

                stores.Add(store);
            }

            ViewBag.wono = wono;
            ViewBag.order_cust = order_cust;
            ViewBag.engsr = engsr;
            ViewBag.inputdate_start = inputdate_start;
            ViewBag.inputdate_end = inputdate_end;

            // 查詢後先依 eng_sr 排序，再分頁，確保每頁筆數正確
            var sortedStores = stores.OrderBy(s => s.eng_sr).ToList();
            return View("StoreHouseStock", "_LayoutMember", sortedStores.ToPagedList(currentPage, pagesize));
        }
        #endregion
        #region 新增庫存 View
        //create
        /// <summary>
        /// 新增 View
        /// </summary>
        /// <returns></returns>
        public ActionResult StockCreate()
        {
            // 權限不足或未登入
            var authResult = Login_Authentication();
            if (authResult != null)
            {
                return authResult;
            }

            Session["IGroup"] = string.Empty;
            if (string.IsNullOrEmpty(Session["Member"].ToString()))
            {
                return View();
            }
            var model = new E_StoreHouseStock();

            return View("StockCreate", "_LayoutMember", model);
        }
        #endregion
        #region 檢查工單是否存在
        /// <summary>
        /// 檢查E_ZRSD19 / E_ZRSD14P 工單是否存在
        /// </summary>
        /// <param name="input">工單</param>
        /// <returns></returns>
        [HttpPost]
        public JsonResult CheckData(string input)
        {
            // 先查詢 E_ZRSD19 與 E_ZRSD14P，若皆無資料則回傳 not found
            string resultItem = null;

            // WHY: 參數為空值時直接回傳 not found，避免誤判
            if (string.IsNullOrWhiteSpace(input))
            {
                resultItem = null;
            }
            else
            {
                var data19 = db.E_ZRSD19.FirstOrDefault(o => o.wono == input);
                if (data19 != null && !string.IsNullOrWhiteSpace(data19.item))
                {
                    resultItem = data19.item;
                }
                else
                {
                    var data14p = db.E_ZRSD14P.FirstOrDefault(o => o.wono == input);
                    if (data14p != null && !string.IsNullOrWhiteSpace(data14p.item))
                    {
                        resultItem = data14p.item;
                    }
                }
            }

            // WHY: 只設定一次 TempData，避免重複與混淆
            TempData["CheckData"] = string.IsNullOrWhiteSpace(resultItem) ? "not found" : "";

            // WHY: 統一回傳結果，提升可維護性
            return Json(new { result = resultItem ?? "not found" });
        }
        #endregion
        #region 新增庫存
        /// <summary>
        /// 新增資料
        /// </summary>
        /// <param name="wono">工單</param>
        /// <param name="quantity">數量</param>
        /// <param name="position">儲位</param>
        /// <param name="mark">備註</param>
        /// <param name="package">包裝</param>
        /// <param name="buttonName">群組</param>
        /// <param name="isActive">無業單新增勾選/param>
        /// <param name="box_quantity">箱數</param>
        /// <param name="transportation">出貨方式</param>
        /// <param name="nowono">列背景色是否顯示-Y(無單號)</param>
        /// <returns></returns>
        [HttpPost]
        public ActionResult StockCreate(string wono, int? quantity, string position, string mark, string package,
            string buttonName, bool isActive, int? box_quantity, string transportation, string nowono = "N")
        {
            /*// 每日排程"\\192.168.4.11\33 倉庫備料單\07. 收發組\成倉\成倉SAP_Excel\"把EXCEL檔案
             * -[ZRSD19]取得寫入 E_ZRSD19
             * -[ZRSD14P]取得寫入 E_ZRSD14P
             * Update E_StoreHouseStock
             */
            // ZRSD19 / ZRSD14P 查無資料無法新增情況
            if (TempData["CheckData"] != null && !string.IsNullOrWhiteSpace(TempData["CheckData"].ToString()))
            {
                ViewBag.err = "ZRSD19/ZRSD14P 查無資料，無法新增!";
                // 保留原欄位值並顯示錯誤
                var model = new WebStoreHouse.Models.E_StoreHouseStock
                {
                    wono = wono,
                    quantity = quantity,
                    position = position,
                    mark = mark,
                    package = package,
                    box_quantity = box_quantity
                };
                return View("StockCreate", "_LayoutMember", model);
            }

            // 若沒有錯誤，建立新的 IGroup（群組識別）供後續使用
            Session["IGroup"] = Guid.NewGuid().ToString();
            TempData["IGroup"] = Session["IGroup"];
            if (buttonName == "unGroup")
            {
                //移除NewGuid
                Session["IGroup"] = "";
                TempData["IGroup"] = "";
                var model = new WebStoreHouse.Models.E_StoreHouseStock
                {
                    wono = wono,
                    quantity = quantity,
                    position = position,
                    mark = mark,
                    package = package,
                    box_quantity = box_quantity
                };
                return View("StockCreate", "_LayoutMember", model);
            }

            // 封裝驗證重複邏輯
            if (!ValidateStockInput(quantity, package, position, out string errMsg))
            {
                ViewBag.err = errMsg;
                // 保留原欄位值
                var model = new WebStoreHouse.Models.E_StoreHouseStock
                {
                    wono = wono,
                    quantity = quantity,
                    position = position,
                    mark = mark,
                    package = package,
                    box_quantity = box_quantity
                };
                return View("StockCreate", "_LayoutMember", model);
            }

            if (buttonName != "Group" && isActive == false)
            {
                // 有單號新增庫存
                AddStockWithOrder(wono, quantity.Value, position, mark, package, box_quantity, transportation, nowono);

                // 如果按下[再一筆]按鈕則返回新增畫面
                if (buttonName == "Again")
                {
                    ModelState.Clear();
                    return View("StockCreate", "_LayoutMember");
                }
            }

            // 無單號新增庫存
            if (isActive)
            {
                AddStockWithoutOrder(wono, quantity.Value, position, mark, package, box_quantity, transportation);
            }

            // 按下Create按鈕則回到庫存畫面
            return RedirectToAction("StoreHouseStock");
        }

        #region 驗證輸入欄位
        /// <summary>
        /// 驗證庫存輸入欄位
        /// </summary>
        /// <param name="quantity">數量</param>
        /// <param name="package">包裝</param>
        /// <param name="position">儲位</param>
        private bool ValidateStockInput(int? quantity, string package, string position, out string errMsg)
        {
            string errMsgs = string.Empty;
            if (!quantity.HasValue || string.IsNullOrWhiteSpace(quantity.ToString()))
            {
                errMsgs = "數量不可為空!";
            }
            if (string.IsNullOrWhiteSpace(package))
            {
                errMsgs += "包裝不可為空!";
            }
            if (string.IsNullOrWhiteSpace(position))
            {
                errMsgs += "儲位不可為空!";
            }
            if (!string.IsNullOrWhiteSpace(errMsgs))
            {
                errMsg = errMsgs;
                return false;
            }
            errMsg = string.Empty;
            return true;
        }
        #endregion
        #region 有單號新增庫存
        /// <summary>
        /// 有單號新增庫存
        /// </summary>
        /// <param name="wono">工單</param>
        /// <param name="quantity">數量</param>
        /// <param name="position">儲位</param>
        /// <param name="mark">備註</param>
        /// <param name="package">包裝</param>
        /// <param name="isActive"></param>
        /// <param name="box_quantity">箱數</param>
        /// <param name="transportation">出貨方式</param>
        /// <param name="nowono">列背景色是否顯示-Y(無單號)</param>
        private void AddStockWithOrder(string wono, int quantity, string position, string mark, string package, int? box_quantity, string transportation, string nowono)
        {
            #region step 1. 寫入庫存表 In 作業
            var stockIn = new E_StoreHouseStock_In
            {
                nowono = nowono,
                wono = wono,
                quantity = quantity,
                box_quantity = box_quantity,
                position = position,
                mark = mark,
                package = package,
                transportation = transportation,
                // 20250821 因應歷史資料的入庫日期的呈現要完整日期，原本寫入入庫日期只存日期，改為存完整日期+時間 By Jesse
                inputdate = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"),
                Igroup = Session["IGroup"].ToString()
            };
            db.E_StoreHouseStock_In.Add(stockIn);
            db.SaveChanges();
            #endregion
            #region step 2. 寫入庫存總表作業
            var data19 = db.E_ZRSD19.FirstOrDefault(o => o.wono == wono);
            var data14p = db.E_ZRSD14P.FirstOrDefault(o => o.wono == wono);
            #region ZRSD19
            if (data19 != null && !string.IsNullOrEmpty(data19.item))
            {
                // 找到剛剛寫入 E_StoreHouseStock_In 的sno
                /*取得[E_StoreHouseStock_In]最新的入庫記錄時*/
                int sno = db.E_StoreHouseStock_In.Where(x => x.wono == wono).OrderByDescending(x => x.sno).Select(x => x.sno).FirstOrDefault();

                /*// 每日排程"\\192.168.4.11\33 倉庫備料單\07. 收發組\成倉\成倉SAP_Excel\"把EXCEL檔案
                 * -[ZRSD19]取得寫入 E_ZRSD19
                 * -[ZRSD14P]取得寫入 E_Kf10
                 * -[KQ30]取得寫入 E_KQ30
                 * -[KF10]取得寫入 E_Kf10
                 * Update E_StoreHouseStock
                 */
                // 透過這筆sno的wono資料查詢相關資訊
                var query = from a in db.E_ZRSD19
                            join b in (from E_Kf10 in db.E_Kf10 group E_Kf10 by E_Kf10.Material into g select new { Material = g.Key, Unrestricted = (int?)g.Sum(p => p.Unrestricted) })
                                on a.item equals b.Material into b_join
                            from b in b_join.DefaultIfEmpty()
                            join c in (from E_KQ30 in db.E_KQ30 group E_KQ30 by E_KQ30.Material into g select new { Material = g.Key, Unrestricted = (int?)g.Sum(p => p.Unrestricted) })
                                on a.item equals c.Material into c_join
                            from c in c_join.DefaultIfEmpty()
                            where a.wono == wono
                            select new
                            {
                                a.wono_cust,
                                a.item,
                                a.order_quantity,
                                kf10 = b.Unrestricted ?? 0,
                                kq30 = c.Unrestricted ?? 0,
                                a.wono_inStoreCount,
                                a.shipped_quantity,
                                a.unshipped_quantity,
                                a.borrow_count,
                                a.due_date,
                            };

                //將該筆wono資料寫入庫存表
                //寫入前確認是否重複 ; stockCount > 0 代表庫存數++
                var stockCount = db.E_StoreHouseStock.Where(stock => stock.wono == wono && stock.position == position && stock.del_flag != "D");
                if (stockCount.Any())
                {
                    var tstock = stockCount.FirstOrDefault();
                    tstock.quantity += quantity;
                    tstock.box_quantity += box_quantity;
                }
                else
                {
                    foreach (var q in query)
                    {
                        var houseStock = new E_StoreHouseStock
                        {
                            sno = sno,
                            nowono = nowono,
                            wono = wono,
                            cust_wono = q.wono_cust,
                            eng_sr = q.item,
                            order_count = q.order_quantity,
                            quantity = quantity,
                            box_quantity = box_quantity,
                            kf10 = q.kf10,
                            kq30 = q.kq30,
                            position = position,
                            acc_in = q.wono_inStoreCount,//庫存
                            outed = q.shipped_quantity,
                            notout = q.unshipped_quantity,
                            borrow = q.borrow_count,
                            due_date = q.due_date,
                            mark = mark,
                            // 20250821 因應歷史資料的入庫日期的呈現要完整日期，原本寫入入庫日期只存日期，改為存完整日期+時間 By Jesse
                            inputdate = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"),
                            package = package,
                            transportation = transportation,
                            Igroup = Session["IGroup"].ToString()
                        };
                        db.E_StoreHouseStock.Add(houseStock);
                    }
                }
                try
                {
                    db.SaveChanges();
                }
                catch (Exception)
                {
                    // TODO: 記錄錯誤日誌
                }
            }
            #endregion
            #region ZRSD14P
            else if (data14p != null && !string.IsNullOrEmpty(data14p.item))
            {
                /*取得[E_StoreHouseStock_In]最新的入庫記錄時*/
                var chk_in = db.E_StoreHouseStock_In.Where(_in => _in.wono == wono).OrderByDescending(_in => _in.sno).FirstOrDefault();
                var houseStock = new E_StoreHouseStock
                {
                    sno = chk_in.sno,
                    nowono = chk_in.nowono,
                    wono = chk_in.wono,
                    quantity = chk_in.quantity,
                    box_quantity = chk_in.box_quantity,
                    position = chk_in.position,
                    mark = chk_in.mark,
                    // 20250821 因應歷史資料的入庫日期的呈現要完整日期，原本寫入入庫日期只存日期，改為存完整日期+時間 By Jesse
                    inputdate = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"),
                    package = chk_in.package,
                    transportation = chk_in.transportation,
                    order_count = 0,
                    kf10 = 0,
                    kq30 = 0,
                    acc_in = 0,
                    outed = 0,
                    notout = 0,
                    borrow = 0,
                    due_date = DateTime.Now,
                };
                db.E_StoreHouseStock.Add(houseStock);
                try
                {
                    db.SaveChanges();
                }
                catch (Exception)
                {
                    // TODO: 記錄錯誤日誌
                }
            }
            #endregion
            #endregion
        }
        #endregion
        #region 無單號新增庫存
        /// <summary>
        /// 無單號新增庫存
        /// </summary>
        /// <param name="wono">工單</param>
        /// <param name="quantity">數量</param>
        /// <param name="position">儲位</param>
        /// <param name="mark">備註</param>
        /// <param name="package">包裝</param>
        /// <param name="isActive"></param>
        /// <param name="box_quantity">箱數</param>
        /// <param name="transportation">出貨方式</param>
        private void AddStockWithoutOrder(string wono, int quantity, string position, string mark, string package, int? box_quantity, string transportation)
        {
            #region step 1. 寫入庫存表 In 作業
            var stockIn = new E_StoreHouseStock_In
            {
                nowono = "Y",//列背景色是否顯示-Y(無單號)
                wono = wono,
                quantity = quantity,
                box_quantity = box_quantity,
                position = position,
                mark = mark,
                package = package,
                transportation = transportation,
                // 20250821 因應歷史資料的入庫日期的呈現要完整日期，原本寫入入庫日期只存日期，改為存完整日期+時間 By Jesse
                inputdate = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
            };
            try
            {
                db.E_StoreHouseStock_In.Add(stockIn);
                db.SaveChanges();
            }
            catch (Exception)
            {
                // TODO: 記錄錯誤日誌
                throw;
            }
            #endregion
            #region step 2. 寫入庫存總表作業
            /*取得[E_StoreHouseStock_In]最新的入庫記錄時*/
            var chk_in = db.E_StoreHouseStock_In.Where(_in => _in.wono == wono).OrderByDescending(_in => _in.sno).FirstOrDefault();
            var houseStock = new E_StoreHouseStock
            {
                sno = chk_in.sno,
                nowono = chk_in.nowono,
                wono = chk_in.wono,
                quantity = chk_in.quantity,
                box_quantity = chk_in.box_quantity,
                position = chk_in.position,
                mark = chk_in.mark,
                // 20250821 因應歷史資料的入庫日期的呈現要完整日期，原本寫入入庫日期只存日期，改為存完整日期+時間 By Jesse
                inputdate = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"),
                package = chk_in.package,
                transportation = chk_in.transportation,
                order_count = 0,
                kf10 = 0,
                kq30 = 0,
                acc_in = 0,
                outed = 0,
                notout = 0,
                borrow = 0,
                due_date = DateTime.Now,
            };
            db.E_StoreHouseStock.Add(houseStock);
            try
            {
                db.SaveChanges();
            }
            catch (Exception)
            {
                // TODO: 記錄錯誤日誌
            }
            #endregion
        }
        #endregion

        #endregion
        #region 修改庫存
        /// <summary>
        /// 修改庫存資料 View
        /// </summary>
        /// <param name="serialno"></param>
        /// <returns></returns>
        public ActionResult EditStock(int serialno)
        {
            // 權限不足或未登入
            var authResult = Login_Authentication();
            if (authResult != null)
            {
                return authResult;
            }

            var result = db.E_StoreHouseStock.SingleOrDefault(m => m.serialno == serialno);
            return View("EditStock", "_LayoutMember", result);
        }
        /// <summary>
        /// 修改庫存資料
        /// </summary>
        /// <param name="serialno"></param>
        /// <param name="position"></param>
        /// <param name="quantity"></param>
        /// <param name="mark"></param>
        /// <param name="wono"></param>
        /// <param name="oldP"></param>
        /// <param name="oldM"></param>
        /// <returns></returns>
        [HttpPost]
        public ActionResult EditStock(int serialno, string position, int quantity, string mark, string wono, string oldP, string oldM)
        {
            #region (原)找出舊資料將其複製並備註轉庫位
            /*
             //找出wono.position為條件的資料, 將serialno抓出來的數量資料累加進以position為基底的數量
            var result = db.E_StoreHouseStock
                .Where(m => m.wono == wono && m.position == oldP && m.del_flag != "D").FirstOrDefault();
            var onlyMarkUpdated = true;
            if (result != null)
            {
                result.quantity += quantity;
                result.mark = mark;
                result.position = position;
                db.SaveChanges();

                // 判断是否只修改了mark
                var newMark = result.mark;
                var newPosition = result.position;
                onlyMarkUpdated = oldM != newMark && oldP != newPosition;

                if (!onlyMarkUpdated)
                {
                    // 只修改了mark，需要還原quantity和mark值
                    result.quantity -= quantity;
                    //result.mark = oldMark;
                    db.SaveChanges();
                }
            }
            //最後將原始(serialno)位置庫存設定為0 設定del_flag='D' ,設定為D不備搜尋出來,del備註自動填寫轉庫位
            var sourceData = db.E_StoreHouseStock.SingleOrDefault(m => m.serialno == serialno);

            if (sourceData != null && onlyMarkUpdated )
            {
                sourceData.del_flag = "D";
                sourceData.quantity = 0;
                sourceData.mark = "轉庫位";
                db.SaveChanges();
            }
             */
            #endregion
            //找出New資料
            var newData = db.E_StoreHouseStock
                .Where(s => s.wono == wono && s.position == position && s.del_flag != "D").FirstOrDefault();
            if (newData != null)//有資料
            {
                newData.position = position;
                newData.quantity += quantity;
                newData.mark = mark;

                #region 原始資料
                //最後將原始(serialno)位置庫存設定為0 設定del_flag='D' ,設定為D不備搜尋出來,del備註自動填寫轉庫位
                var sourceData = db.E_StoreHouseStock.SingleOrDefault(m => m.serialno == serialno);
                sourceData.del_flag = "D";
                sourceData.quantity = 0;
                sourceData.mark = "轉庫位";
                db.SaveChanges();
                #endregion
            }
            else
            {
                var oldData = db.E_StoreHouseStock
                .Where(s => s.wono == wono && s.position == oldP && s.del_flag != "D").FirstOrDefault();
                oldData.position = position;
                oldData.mark = mark;
                db.SaveChanges();
            }


            return RedirectToAction("StoreHouseStock");
        }

        /// <summary>
        /// 刪除庫存資料（軟刪除，僅設置 del_flag 為 'D'，不會從資料庫實體刪除）。
        /// </summary>
        /// <param name="serialno">庫存資料的唯一識別碼（serialno）。</param>
        /// <param name="igroup">群組代碼（Igroup），若有指定則刪除該群組下所有庫存資料。</param>
        /// <returns>
        /// 刪除後重新導向至 <see cref="StoreHouseStock"/> 庫存查詢頁面。
        /// 若未登入則導向登入頁。
        /// </returns>
        public ActionResult DeleteStock(int? serialno, string igroup)
        {
            // 權限不足或未登入
            var authResult = Login_Authentication();
            if (authResult != null)
            {
                return authResult;
            }

            // 刪除非真正刪除資料，僅設置 del_flag 以後查詢時不顯示
            List<E_StoreHouseStock> stock;
            if (!string.IsNullOrEmpty(igroup))
            {
                // 若有指定群組，則刪除該群組下所有庫存資料
                stock = db.E_StoreHouseStock.Where(m => m.Igroup == igroup).ToList();
            }
            else
            {
                // 否則僅刪除指定 serialno 的單筆庫存資料
                stock = db.E_StoreHouseStock.Where(m => m.serialno == serialno).ToList();
            }

            foreach (var stocks in stock)
            {
                stocks.del_flag = "D";
            }

            db.SaveChanges();
            return RedirectToAction("StoreHouseStock");
        }
        #endregion
        #region  匯出倉庫庫存 Excel
        /// <summary>
        /// 匯出倉庫庫存 Excel
        /// </summary>
        /// <param name="wono">工單號碼</param>
        /// <param name="order_cust">客戶業單</param>
        /// <param name="engsr">機種名稱</param>
        /// <param name="inputdate_start">入庫日期起</param>
        /// <param name="inputdate_end">入庫日期迄</param>
        /// <returns>Excel 檔案下載</returns>
        [HttpGet]
        public ActionResult ExportStoreHouseStock(string wono, string order_cust, string engsr, string inputdate_start, string inputdate_end)
        {
            // 登入驗證
            var authResult = Login_Authentication();
            if (authResult != null)
            {
                Response.StatusCode = 401;
                return Content("未登入，請重新登入。", "text/plain");
            }

            // 20250821 因應歷史資料的入庫日期的呈現要完整日期，原本寫入入庫日期只存日期，改為存完整日期+時間，入庫日期修改顯示日期 By Jesse
            // 查詢條件與 SQL 與 StoreHouseStock 相同
            List<SqlParameter> parmL = new List<SqlParameter>();
            string strsql = @"select serialno,sno,nowono,wono,cust_wono,eng_sr,order_count,quantity,box_quantity,kf10,kq30,
                            sap_in,position,acc_in,outed,notout,borrow,due_date,mark,CONVERT(varchar(10), inputdate, 120) AS inputdate,package,output_local,Igroup 
                            from E_StoreHouseStock where quantity >0 and (del_flag is null or del_flag <> 'D') and (igroup ='' or Igroup is null) ";
            if (!string.IsNullOrWhiteSpace(wono))
            {
                strsql += " and wono = @wono ";
                parmL.Add(new SqlParameter("wono", wono));
            }
            if (!string.IsNullOrWhiteSpace(order_cust))
            {
                strsql += " and cust_wono = @order_cust ";
                parmL.Add(new SqlParameter("order_cust", order_cust));
            }
            if (!string.IsNullOrWhiteSpace(engsr))
            {
                strsql += " and eng_sr = @engsr ";
                parmL.Add(new SqlParameter("engsr", engsr));
            }
            // 20250821 因應歷史資料的入庫日期的呈現要完整日期，原本寫入入庫日期只存日期，改為存完整日期+時間，入庫日期修改顯示日期 By Jesse
            if (!string.IsNullOrWhiteSpace(inputdate_start))
            {
                strsql += " and CONVERT(varchar(10), inputdate, 120) >= @inputdate_start ";
                parmL.Add(new SqlParameter("inputdate_start", inputdate_start));
            }
            if (!string.IsNullOrWhiteSpace(inputdate_end))
            {
                strsql += " and CONVERT(varchar(10), inputdate, 120) <= @inputdate_end ";
                parmL.Add(new SqlParameter("inputdate_end", inputdate_end));
            }

            // 20250821 因應歷史資料的入庫日期的呈現要完整日期，原本寫入入庫日期只存日期，改為存完整日期+時間，入庫日期修改顯示日期 By Jesse
            strsql += @" union
                        select serialno,sno,nowono,wono,cust_wono,eng_sr,order_count,quantity,box_quantity,kf10,kq30,
                        sap_in,position,acc_in,outed,notout,borrow,due_date,mark,CONVERT(varchar(10), inputdate, 120) AS inputdate,package,output_local,Igroup 
                        from (Select *,ROW_NUMBER() Over (Partition By Igroup Order By CONVERT(varchar(10), inputdate, 120) Desc) As Sort 
                        From E_StoreHouseStock where igroup <>'' and igroup is not null and (del_flag is null or del_flag <> 'D') 
                        and quantity >0) s where sort=1 ";
            if (!string.IsNullOrWhiteSpace(wono))
            {
                strsql += " and wono = @wono2 ";
                parmL.Add(new SqlParameter("wono2", wono));
            }
            if (!string.IsNullOrWhiteSpace(order_cust))
            {
                strsql += " and cust_wono = @order_cust2 ";
                parmL.Add(new SqlParameter("order_cust2", order_cust));
            }
            if (!string.IsNullOrWhiteSpace(engsr))
            {
                strsql += " and eng_sr = @engsr2 ";
                parmL.Add(new SqlParameter("engsr2", engsr));
            }
            // 20250821 因應歷史資料的入庫日期的呈現要完整日期，原本寫入入庫日期只存日期，改為存完整日期+時間，入庫日期修改顯示日期 By Jesse
            if (!string.IsNullOrWhiteSpace(inputdate_start))
            {
                strsql += " and CONVERT(varchar(10), inputdate, 120) >= @inputdate_start2 ";
                parmL.Add(new SqlParameter("inputdate_start2", inputdate_start));
            }
            if (!string.IsNullOrWhiteSpace(inputdate_end))
            {
                strsql += " and CONVERT(varchar(10), inputdate, 120) <= @inputdate_end2 ";
                parmL.Add(new SqlParameter("inputdate_end2", inputdate_end));
            }

            strsql += @" order by eng_sr ";

            var dr = dbMethod.ExecuteReaderPmsList(strsql, CommandType.Text, parmL);
            // 檢查是否有資料
            if (dr == null || !dr.HasRows)
            {
                // 沒有資料時回傳提示訊息或空檔案
                TempData["ErrorMessage"] = "查無資料，無法匯出 Excel。";
                Response.StatusCode = 404;
                return Content("查無資料，無法匯出 Excel。", "text/plain");
                //return RedirectToAction("StoreHouseStock");
            }

            var data = new List<StoreHouseStock>();
            while (dr.Read())
            {
                var store = new StoreHouseStock();
                store.serialno = (int)dr["serialno"];
                store.nowono = dr["nowono"].ToString();
                store.sno = (int)dr["sno"];
                store.wono = dr["wono"].ToString();
                store.cust_wono = dr["cust_wono"].ToString();
                store.eng_sr = dr["eng_sr"].ToString();
                store.order_count = (int)dr["order_count"];
                store.quantity = (int)dr["quantity"];
                store.box_quantity = (int)dr["box_quantity"];
                store.kf10 = (int)dr["kf10"];
                store.kq30 = (int)dr["kq30"];
                store.sap_in = dr["sap_in"].ToString();
                store.position = dr["position"].ToString();
                store.acc_in = (int)dr["acc_in"];
                store.outed = (int)dr["outed"];
                store.notout = (int)dr["notout"];
                store.borrow = (int)dr["borrow"];
                store.due_date = (DateTime)dr["due_date"];
                store.mark = dr["mark"].ToString();
                store.inputdate = dr["inputdate"].ToString();
                store.package = dr["package"].ToString();
                store.output_local = dr["output_local"].ToString();
                store.Igroup = dr["Igroup"].ToString();
                data.Add(store);
            }
            dr.Close();

            // 產生 Excel
            IWorkbook workbook = new XSSFWorkbook();
            ISheet sheet = workbook.CreateSheet("成倉庫存");
            // 標題列
            var header = new[] { "工單", "客戶業單", "機種名稱", "訂單數量", "箱數", "數量", "KF10", "KQ30", "儲位", "累積入庫", "預計銷單日", "備註", "入庫日期", "包裝" };
            IRow rowHeader = sheet.CreateRow(0);
            for (int i = 0; i < header.Length; i++)
            {
                rowHeader.CreateCell(i).SetCellValue(header[i]);
            }
            // 資料列
            int rowIdx = 1;
            foreach (var s in data)
            {
                IRow row = sheet.CreateRow(rowIdx++);
                row.CreateCell(0).SetCellValue(s.wono);
                row.CreateCell(1).SetCellValue(s.cust_wono);
                row.CreateCell(2).SetCellValue(s.eng_sr);
                row.CreateCell(3).SetCellValue(s.order_count);
                row.CreateCell(4).SetCellValue(s.box_quantity);
                row.CreateCell(5).SetCellValue(s.quantity);
                row.CreateCell(6).SetCellValue(s.kf10);
                row.CreateCell(7).SetCellValue(s.kq30);
                row.CreateCell(8).SetCellValue(s.position);
                row.CreateCell(9).SetCellValue(s.acc_in);
                row.CreateCell(10).SetCellValue(((DateTime)s.due_date).ToString("yyyy-MM-dd"));
                row.CreateCell(11).SetCellValue(s.mark);
                row.CreateCell(12).SetCellValue(s.inputdate);
                row.CreateCell(13).SetCellValue(s.package);
            }
            // 自動欄寬
            for (int i = 0; i < header.Length; i++)
            {
                sheet.AutoSizeColumn(i);
            }
            using (var exportData = new MemoryStream())
            {
                workbook.Write(exportData);
                string fileName = $"倉庫庫存_{DateTime.Now:yyyyMMddHHmmss}.xlsx";
                // RFC 5987 percent-encode (UTF-8)
                string encodedFileName = System.Web.HttpUtility.UrlEncode(fileName, System.Text.Encoding.UTF8).Replace("+", "%20");
                Response.Clear();
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                // 同時支援 filename 與 filename* (RFC 5987)
                Response.AddHeader("Content-Disposition", $"attachment; filename=\"{fileName}\"; filename*=UTF-8''{encodedFileName}");
                Response.BinaryWrite(exportData.ToArray());
                Response.End();
                return new EmptyResult();
            }
        }
        #endregion
        #endregion

        #endregion

        #region 入庫資料查詢
        #region 查詢
        /// <summary>
        /// 入庫資料查詢
        /// </summary>
        /// <param name="inputdate_start">入庫日期(起)</param>
        /// <param name="wono">工單</param>
        /// <param name="engsr">機種</param>
        /// <param name="page">分頁</param>
        /// <returns></returns>
        public ActionResult StoreHouseStockEntryQuery(string inputdate_start, string wono, string engsr, int page = 1)
        {
            // WHY: 改為回傳分頁 JSON，符合前端 fetch 需求
            try
            {
                var authResult = Login_Authentication();
                if (authResult != null)
                {
                    // 未登入直接回傳 JSON 錯誤
                    return Json(new { success = false, error = "未登入，請重新登入。" }, JsonRequestBehavior.AllowGet);
                }
                List<SqlParameter> parmL = new List<SqlParameter>();
                string strsql = @"SELECT a.quantity, a.wono, a.inputdate, b.eng_sr 
                                  FROM E_StoreHouseStock_In a 
                                  INNER JOIN E_StoreHouseStock b ON b.wono = a.wono AND b.eng_sr IS NOT NULL AND b.eng_sr <> '' AND a.wono IS NOT NULL AND a.wono <> ''  
                                  WHERE 1=1 ";
                // 擇一查詢條件
                if (!string.IsNullOrWhiteSpace(inputdate_start))
                {
                    strsql += " and CONVERT(varchar(10), a.inputdate, 120) = @inputdate_start ";
                    parmL.Add(new SqlParameter("inputdate_start", inputdate_start));
                }
                if (!string.IsNullOrWhiteSpace(wono))
                {
                    strsql += " and a.wono = @wono ";
                    parmL.Add(new SqlParameter("wono", wono));
                }
                if (!string.IsNullOrWhiteSpace(engsr))
                {
                    strsql += " and b.eng_sr = @engsr ";
                    parmL.Add(new SqlParameter("engsr", engsr));
                }

                strsql += " GROUP BY a.quantity, a.wono, a.inputdate, b.eng_sr order by a.inputdate ";

                // 先查詢全部資料，計算總筆數
                SqlDataReader dr = dbMethod.ExecuteReaderPmsList(strsql, CommandType.Text, parmL);
                var allList = new List<object>();
                while (dr.Read())
                {
                    DateTime inputDate;
                    DateTime.TryParse(dr["inputdate"].ToString(), out inputDate);
                    int qty = 0;
                    int.TryParse(dr["quantity"].ToString(), out qty);
                    allList.Add(new
                    {
                        InputDate = inputDate.ToString("yyyy-MM-dd HH:mm:ss"),
                        WorkOrder = dr["wono"].ToString(),
                        Model = dr["eng_sr"].ToString(),
                        Quantity = qty
                    });
                }
                dr.Close();
                int totalCount = allList.Count;
                int pageSize = entryPageSize > 0 ? entryPageSize : 30;
                int currentPage = page < 1 ? 1 : page;
                var pagedList = allList.Skip((currentPage - 1) * pageSize).Take(pageSize).ToList();
                // WHY: 回傳分頁資料與總筆數
                return Json(new { data = pagedList, total = totalCount }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                // 回傳 JSON 格式錯誤訊息
                return Json(new { success = false, error = ex.Message }, JsonRequestBehavior.AllowGet);
            }
        }
        #endregion
        #region 入庫資料匯出 Excel
        /// <summary>
        /// 入庫資料匯出 Excel
        /// </summary>
        [HttpGet]
        public ActionResult StoreHouseStockEntryExport(string inputdate_start, string wono, string engsr)
        {
            try
            {
                var authResult = Login_Authentication();
                if (authResult != null)
                {
                    Response.StatusCode = 401;
                    return Content("未登入，請重新登入。", "text/plain");
                }

                List<SqlParameter> parmL = new List<SqlParameter>();

                string strsql = @"SELECT a.quantity, a.wono, a.inputdate, b.eng_sr 
                                  FROM E_StoreHouseStock_In a 
                                  INNER JOIN E_StoreHouseStock b ON b.wono = a.wono AND b.eng_sr IS NOT NULL AND b.eng_sr <> '' AND a.wono IS NOT NULL AND a.wono <> ''
                                  WHERE 1=1 ";

                if (!string.IsNullOrWhiteSpace(inputdate_start))
                {
                    strsql += " and CONVERT(varchar(10), a.inputdate, 120) = @inputdate_start ";
                    parmL.Add(new SqlParameter("inputdate_start", inputdate_start));
                }
                if (!string.IsNullOrWhiteSpace(wono))
                {
                    strsql += " and a.wono = @wono ";
                    parmL.Add(new SqlParameter("wono", wono));
                }
                if (!string.IsNullOrWhiteSpace(engsr))
                {
                    strsql += " and a.eng_sr = @engsr ";
                    parmL.Add(new SqlParameter("engsr", engsr));
                }

                strsql += " GROUP BY a.quantity, a.wono, a.inputdate, b.eng_sr order by a.inputdate ";

                SqlDataReader dr = dbMethod.ExecuteReaderPmsList(strsql, CommandType.Text, parmL);
                var data = new List<dynamic>();
                while (dr.Read())
                {
                    DateTime inputDate;
                    DateTime.TryParse(dr["inputdate"].ToString(), out inputDate);
                    int qty = 0;
                    int.TryParse(dr["quantity"].ToString(), out qty);
                    data.Add(new
                    {
                        // 20250821 因應歷史資料的入庫日期的呈現要完整日期，原本寫入入庫日期只存日期，改為存完整日期+時間，入庫日期修改顯示日期 By Jesse
                        InputDate = inputDate.ToString("yyyy-MM-dd HH:mm:ss"),
                        WorkOrder = dr["wono"].ToString(),
                        Model = dr["eng_sr"].ToString(),
                        Quantity = qty
                    });
                }
                dr.Close();

                // 產生 Excel
                IWorkbook workbook = new XSSFWorkbook();
                ISheet sheet = workbook.CreateSheet("入庫資料");
                var header = new[] { "入庫日期", "工單", "機種", "數量" };
                IRow rowHeader = sheet.CreateRow(0);
                for (int i = 0; i < header.Length; i++)
                {
                    rowHeader.CreateCell(i).SetCellValue(header[i]);
                }
                int rowIdx = 1;
                foreach (var s in data)
                {
                    IRow row = sheet.CreateRow(rowIdx++);
                    row.CreateCell(0).SetCellValue(s.InputDate);
                    row.CreateCell(1).SetCellValue(s.WorkOrder);
                    row.CreateCell(2).SetCellValue(s.Model);
                    row.CreateCell(3).SetCellValue(s.Quantity);
                }
                for (int i = 0; i < header.Length; i++)
                {
                    sheet.AutoSizeColumn(i);
                }
                using (var exportData = new MemoryStream())
                {
                    workbook.Write(exportData);
                    string fileName = $"入庫資料_{DateTime.Now:yyyyMMddHHmmss}.xlsx";
                    // 直接於 using block 內取得 byte[]，避免 stream 被關閉
                    return File(exportData.ToArray(),
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        fileName);
                }
            }
            catch (Exception ex)
            {
                return Content("匯出失敗: " + ex.Message, "text/plain");
            }
        }
        #endregion
        #endregion

        #region 出庫資料查詢
        #region 查詢
        /// <summary>
        /// 出庫資料查詢
        /// </summary>
        /// <param name="outputdate_start">出庫日期</param>
        /// <param name="wono">工單</param>
        /// <param name="engsr">機種</param>
        /// <param name="page">分頁</param>
        /// <returns>分頁 JSON</returns>
        public ActionResult StoreHouseStockOutQuery(string outputdate_start, string wono, string engsr, int page = 1)
        {
            // WHY: 回傳分頁 JSON，符合前端 fetch 需求
            try
            {
                var authResult = Login_Authentication();
                if (authResult != null)
                {
                    // 未登入直接回傳 JSON 錯誤
                    return Json(new { success = false, error = "未登入，請重新登入。" }, JsonRequestBehavior.AllowGet);
                }
                List<SqlParameter> parmL = new List<SqlParameter>();
                string strsql = @"SELECT a.chk_date, a.wono, a.eng_sr, a.exp_shipquantity 
                                  FROM E_StoreHouseStock_Order a 
                                  WHERE ISNULL(a.del_flag,'')!='D' ";
                // 擇一查詢條件
                if (!string.IsNullOrWhiteSpace(outputdate_start))
                {
                    strsql += " and CONVERT(varchar(10), a.chk_date, 120) = @outputdate ";
                    parmL.Add(new SqlParameter("outputdate", outputdate_start));
                }
                if (!string.IsNullOrWhiteSpace(wono))
                {
                    strsql += " and a.wono = @wono ";
                    parmL.Add(new SqlParameter("wono", wono));
                }
                if (!string.IsNullOrWhiteSpace(engsr))
                {
                    strsql += " and a.eng_sr = @engsr ";
                    parmL.Add(new SqlParameter("engsr", engsr));
                }

                strsql += " order by a.chk_date ";

                // 先查詢全部資料，計算總筆數
                SqlDataReader dr = dbMethod.ExecuteReaderPmsList(strsql, CommandType.Text, parmL);
                var allList = new List<object>();
                while (dr.Read())
                {
                    DateTime outputDate;
                    DateTime.TryParse(dr["chk_date"].ToString(), out outputDate);
                    int qty = 0;
                    int.TryParse(dr["exp_shipquantity"].ToString(), out qty);
                    allList.Add(new
                    {
                        OutputDate = outputDate.ToString("yyyy-MM-dd"),
                        WorkOrder = dr["wono"].ToString(),
                        Model = dr["eng_sr"].ToString(),
                        Quantity = qty
                    });
                }
                dr.Close();
                int totalCount = allList.Count;
                int pageSize = entryPageSize > 0 ? entryPageSize : 30;
                int currentPage = page < 1 ? 1 : page;
                var pagedList = allList.Skip((currentPage - 1) * pageSize).Take(pageSize).ToList();
                // WHY: 回傳分頁資料與總筆數
                return Json(new { data = pagedList, total = totalCount }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                // 回傳 JSON 格式錯誤訊息
                return Json(new { success = false, error = ex.Message }, JsonRequestBehavior.AllowGet);
            }
        }
        #endregion
        #region 出庫資料匯出 Excel
        /// <summary>
        /// 出庫資料匯出 Excel
        /// </summary>
        /// <param name="outputdate_start">出庫日期</param>
        /// <param name="wono">工單</param>
        /// <param name="engsr">機種</param>
        /// <param name="page">分頁</param>
        /// <returns></returns>
        [HttpGet]
        public ActionResult StoreHouseStockOutExport(string outputdate_start, string wono, string engsr)
        {
            try
            {
                var authResult = Login_Authentication();
                if (authResult != null)
                {
                    Response.StatusCode = 401;
                    return Content("未登入，請重新登入。", "text/plain");
                }

                List<SqlParameter> parmL = new List<SqlParameter>();
                string strsql = @"SELECT a.chk_date, a.wono, a.eng_sr, a.exp_shipquantity 
                                  FROM E_StoreHouseStock_Order a                           
                                  WHERE ISNULL(a.del_flag,'')!='D' ";

                if (!string.IsNullOrWhiteSpace(outputdate_start))
                {
                    strsql += " and CONVERT(varchar(10), a.chk_date, 120) = @outputdate ";
                    parmL.Add(new SqlParameter("outputdate", outputdate_start));
                }
                if (!string.IsNullOrWhiteSpace(wono))
                {
                    strsql += " and a.wono = @wono ";
                    parmL.Add(new SqlParameter("wono", wono));
                }
                if (!string.IsNullOrWhiteSpace(engsr))
                {
                    strsql += " and a.eng_sr = @engsr ";
                    parmL.Add(new SqlParameter("engsr", engsr));
                }

                strsql += " order by a.chk_date ";

                SqlDataReader dr = dbMethod.ExecuteReaderPmsList(strsql, CommandType.Text, parmL);
                var data = new List<dynamic>();
                while (dr.Read())
                {
                    DateTime outputDate;
                    DateTime.TryParse(dr["chk_date"].ToString(), out outputDate);
                    int qty = 0;
                    int.TryParse(dr["exp_shipquantity"].ToString(), out qty);
                    data.Add(new
                    {
                        OutputDate = outputDate.ToString("yyyy-MM-dd"),
                        WorkOrder = dr["wono"].ToString(),
                        Model = dr["eng_sr"].ToString(),
                        Quantity = qty
                    });
                }
                dr.Close();

                // 產生 Excel
                IWorkbook workbook = new XSSFWorkbook();
                ISheet sheet = workbook.CreateSheet("出庫資料");
                var header = new[] { "出庫日期", "工單", "機種", "數量" };
                IRow rowHeader = sheet.CreateRow(0);
                for (int i = 0; i < header.Length; i++)
                {
                    rowHeader.CreateCell(i).SetCellValue(header[i]);
                }
                int rowIdx = 1;
                foreach (var s in data)
                {
                    IRow row = sheet.CreateRow(rowIdx++);
                    row.CreateCell(0).SetCellValue(s.OutputDate);
                    row.CreateCell(1).SetCellValue(s.WorkOrder);
                    row.CreateCell(2).SetCellValue(s.Model);
                    row.CreateCell(3).SetCellValue(s.Quantity);
                }
                for (int i = 0; i < header.Length; i++)
                {
                    sheet.AutoSizeColumn(i);
                }
                using (var exportData = new MemoryStream())
                {
                    workbook.Write(exportData);
                    string fileName = $"出庫資料_{DateTime.Now:yyyyMMddHHmmss}.xlsx";
                    // RFC 5987 percent-encode (UTF-8)
                    string encodedFileName = System.Web.HttpUtility.UrlEncode(fileName, System.Text.Encoding.UTF8).Replace("+", "%20");
                    Response.Clear();
                    Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                    Response.AddHeader("Content-Disposition", $"attachment; filename=\"{fileName}\"; filename*=UTF-8''{encodedFileName}");
                    Response.BinaryWrite(exportData.ToArray());
                    Response.End();
                    return new EmptyResult();
                }
            }
            catch (Exception ex)
            {
                return Content("匯出失敗: " + ex.Message, "text/plain");
            }
        }
        #endregion
        #endregion

        #region 購物車相關功能
        /// <summary>
        /// 購物車
        /// </summary>
        /// <param name="sno"></param>
        /// <returns></returns>
        public ActionResult AddCar(int sno)
        {
            // 權限不足或未登入
            var authResult = Login_Authentication();
            if (authResult != null)
            {
                return authResult;
            }

            //取得會員帳號指定給fUserId
            string UserId = ViewBag.UserId;


            //找出會員放入訂單明細(table)的產品,該產品的fIsApproved為"否"
            //表示該產品是購物車的"狀態"
            //var currentCar = db.E_StoreHouseStock_SC.Where(m => m.sno == sno && m.IsApproved == "N" ).FirstOrDefault();
            //var dataToInsert = from a in db.E_ZRSD13
            //                   select new E_StoreHouseStock_SC
            //                   {
            //                       exp_shipdate = a.gi_date,
            //                       sales_order = a.sales_order,
            //                       cust_wono = a.cust_wono,
            //                       order_item = a.order_item,
            //                       eng_sr = a.item,
            //                       exp_shipquantity = a.quantity,
            //                       sap_mark = a.ship_mark,
            //                       UserId = UserId,
            //                       IsApproved = "N"
            //                   };
            //db.E_StoreHouseStock_SC.AddRange(dataToInsert);

            #region 此段簡化
            /*
            //若currentCar等於null, 表示會員選購的產品不是購物車的"狀態"
            if (currentCar == null)
            {
                //找出目前選購的產品並指定給product
                var Stock = db.E_StoreHouseStock.Where(m => m.sno == sno).FirstOrDefault();
                //找出該筆資料是否為群組;群組代碼為
                string igroup = (from SHS in db.E_StoreHouseStock
                                 where SHS.sno == sno
                                 orderby SHS.sno descending
                                 select SHS.Igroup).FirstOrDefault();
                if (!string.IsNullOrEmpty(igroup))
                {
                    var queryE_StoreHouseStock_SC =
                    from SHS in db.E_StoreHouseStock
                    where
                      SHS.Igroup == igroup
                    select new
                    {
                        UserID = UserId,
                        SHS.quantity,
                        SHS.wono,
                        SHS.eng_sr,
                        SHS.cust_wono,
                    };
                    foreach (var q in queryE_StoreHouseStock_SC)
                    {
                        E_StoreHouseStock_SC Stock_SC = new E_StoreHouseStock_SC
                        {
                            UserId = q.UserID,
                            stock_quantity = q.quantity,
                            wono = q.wono,
                            eng_sr = q.eng_sr,
                            cust_wono = q.cust_wono,
                            sales_order = (q.wono).Substring(0, 7),
                            IsApproved = "N"
                        };
                        db.E_StoreHouseStock_SC.Add(Stock_SC);
                    }
                }
                else
                {
                    //將產品放入訂單明細,因為產品的fIsApproved為"否",表示購物車"狀態"
                    E_StoreHouseStock_SC stock_SC = new E_StoreHouseStock_SC();
                    stock_SC.UserId = UserId;
                    stock_SC.stock_quantity = Stock.quantity;
                    stock_SC.wono = Stock.wono;
                    stock_SC.eng_sr = Stock.eng_sr;
                    stock_SC.cust_wono = Stock.cust_wono;
                    stock_SC.sales_order = (Stock.wono).Substring(0, 7);
                    stock_SC.IsApproved = "N";
                    db.E_StoreHouseStock_SC.Add(stock_SC);
                }
            }
            */
            #endregion

            //執行Home控制器的ShoppingCar動作方法
            return RedirectToAction("StoreHouseStock_SC");

        }
        /// <summary>
        /// 刪除預計出貨明細（購物車）資料。
        /// </summary>
        /// <param name="sno">預計出貨明細的唯一識別碼。</param>
        /// <returns>刪除後重新導向至預計出貨明細頁面。</returns>
        public ActionResult Delete_SC(int sno)
        {
            var del_sc = db.E_StoreHouseStock_SC.Find(sno);
            db.E_StoreHouseStock_SC.Remove(del_sc);
            db.SaveChanges();
            return RedirectToAction("StoreHouseStock_SC");
        }
        /// <summary>
        /// 刪除購物車中指定工單號碼的商品。
        /// </summary>
        /// <param name="wono">工單號碼。</param>
        /// <returns>刪除後重新導向至預計出貨明細頁面。</returns>
        public ActionResult DeleteCar(string wono)
        {
            //依照fId找出要刪除的購物車"狀態"的商品
            var stock_SC = db.E_StoreHouseStock_SC.Where(m => m.wono == wono).FirstOrDefault();
            //刪除購物車"狀態"的產品
            db.E_StoreHouseStock_SC.Remove(stock_SC);
            db.SaveChanges();
            //執行Home控制器的ShoppingCar動作方法
            return RedirectToAction("StoreHouseStock_SC");
        }
        #endregion

        #region 預設出貨
        // 每日排程"\\192.168.4.11\33 倉庫備料單\07. 收發組\成倉\成倉SAP_Excel\"把EXCEL檔案-ZRSD13 寫入 E_ZRSD13取得寫入
        //GET:Function/StoreHouseStock_SC
        /// <summary>
        /// 預設出貨
        /// </summary>
        /// <param name="page"></param>
        /// <returns></returns>
        public ActionResult StoreHouseStock_SC(int page = 1)
        {
            int currentPage = page < 1 ? 1 : page;
            //取得會員帳號指定fUserId
            //string UserId = (Session["Member"] as MaintainViewModels).fUserId;
            //string RoleId = (Session["Member"] as MaintainViewModels).ROLE_ID;
            //ViewBag.Member = UserId;
            //ViewBag.RoleId = RoleId;
            var authResult = Login_Authentication();
            if (authResult != null)
            {
                return authResult;
            }
            //找出未成為訂單明細資料,即購物車內容
            string today = DateTime.Now.ToString("yyyy-MM-dd");

            //已銷單庫存出貨作業新增資料到[預設出貨]-E_StoreHouseStock_SC
            var orderDetails = db.E_StoreHouseStock_SC.Where
                (m => m.IsApproved == "N" && m.exp_shipdate == today).ToList();
            //指定ShoppingCar.cshtml套用_LayoutMember.cshtml,View使用orderDetail模型
            return View("StoreHouseStock_SC", "_LayoutMember", orderDetails.ToPagedList(currentPage, pagesize));
        }

        /*
        * 此動作會先新增tOrder訂單主檔,接著將tOrderDetail訂單明細的購物車"狀態"之產品的fApproved屬性設為"是"
        * 表示該筆產品正式成為訂單的產品之一, 訂單處理完成後執行Home/OrderList動作方法切換到訂單顯示作業
        */
        /// <summary>
        /// 建立訂單主檔與訂單明細，將購物車狀態的產品（IsApproved="N"）轉為已確認（IsApproved="Y"），
        /// 並產生唯一訂單識別碼（OrderGuid），完成後導向訂單列表頁。
        /// </summary>
        /// <param name="fReceiver">收件人</param>
        /// <param name="fEmail">收件人電子郵件</param>
        /// <param name="fAddress">收件人地址</param>
        /// <returns>導向訂單列表頁（OrderList）</returns>
        [HttpPost]
        public ActionResult StoreHouseStock_SC(string fReceiver, string fEmail, string fAddress)
        {
            // 權限不足或未登入
            var authResult = Login_Authentication();
            if (authResult != null)
            {
                return authResult;
            }

            //找出會員帳號指定給fUserId
            string UserId = ViewBag.UserId;

            //建立唯一的識別,並指定給guid變數用來當作訂單編號
            //tOrder的fOrderGuid欄位會關連到tOrderDetail的fOrderGuid欄位
            //形成一對的關係,即一筆訂單資料對應多筆訂單明細(Master-Detail)
            string guid = Guid.NewGuid().ToString();

            //建立訂單主檔資料
            E_StoreHouseStock_Order stock_Order = new E_StoreHouseStock_Order();
            stock_Order.OrderGuid = guid;    //訂單識別碼
            stock_Order.UserId = UserId;
            stock_Order.chk_date = DateTime.Now;
            db.E_StoreHouseStock_Order.Add(stock_Order);
            //找出目前會員在訂單明細中是購物車"狀態"的產品
            var carList = db.E_StoreHouseStock_SC.Where(m => m.IsApproved == "N").ToList();
            //將購物車狀態的fIsApproved設定為"是"表示確認訂購產品
            foreach (var item in carList)
            {
                item.OrderGuid = guid;     //進入該迴圈表示訂單成立,加入訂單識別碼
                item.IsApproved = "Y";
            }
            //更新資料庫, 異動tOrder 和 tOrderDetail
            //完成訂單主檔和訂單明細的更新
            try
            {
                db.SaveChanges();
            }
            catch (Exception)
            {
                // 重新拋出例外以保留原行為
                throw;
            }

            //執行Home控制器的OrderList動作方法
            return RedirectToAction("OrderList");
        }

        /// <summary>
        /// 編輯預計出貨 View
        /// </summary>
        /// <param name="_sno"></param>
        /// <param name="_sales_order"></param>
        /// <returns></returns>
        public CVM13_Stock Edit_View(int _sno, string _sales_order)
        {
            // 找出要編輯的E_StoreHouseStock_SC資料
            CVM13_Stock CVM13 = new CVM13_Stock()
            {
                //找出未成為訂單明細資料,即購物車內容
                storeHouseStock_SCs = db.E_StoreHouseStock_SC.Where(m => m.sno == _sno).ToList(),
                //storeHouseStock_SCs = new List<E_StoreHouseStock_SC>() { stockSC },
                storeHouseStocks = db.E_StoreHouseStock
                .Where(m => (m.wono).Substring(0, 7) == _sales_order && m.del_flag != "D" && m.quantity > 0).ToList()
            };
            return CVM13;
        }
        /// <summary>
        /// 編輯預計出貨明細的 View。
        /// </summary>
        /// <param name="sno">預計出貨明細的唯一識別碼。</param>
        /// <param name="sales_order">業務單號。</param>
        /// <returns>
        /// 回傳 Edit_SC 頁面，並帶入指定 sno 與 sales_order 的預計出貨明細資料。
        /// </returns>
        public ActionResult Edit_SC(int sno, string sales_order)
        {
            Edit_View(sno, sales_order);

            return View("Edit_SC", "_LayoutMember", Edit_View(sno, sales_order));
        }

        /// <summary>
        /// 編輯預計出貨明細的資料。
        /// 1. 若 company_code 為空，則清除公司相關欄位。
        /// 2. 若 company_code 不為空，檢查庫存數量是否足夠，並更新出貨明細與公司資訊。
        /// 3. 若公司代號不存在，顯示錯誤訊息。
        /// </summary>
        /// <param name="sno">預計出貨明細的唯一識別碼。</param>
        /// <param name="sales_order">業務單號。</param>
        /// <param name="cust_wono">客戶工單號。</param>
        /// <param name="exp_shipquantity">預計出貨數量。</param>
        /// <param name="sap_mark">SAP 備註。</param>
        /// <param name="stock_quantity">成倉庫存數量。</param>
        /// <param name="sap_in">SAP 入庫。</param>
        /// <param name="company_code">公司代號。</param>
        /// <param name="wono">工單號。</param>
        /// <param name="position">庫位。</param>
        /// <param name="eng_sr">機種名稱。</param>
        /// <param name="ssno">預計出貨明細的 sno。</param>
        /// <param name="box_quantity">箱數。</param>
        /// <param name="transportation">出貨方式。</param>
        /// <returns>編輯完成後導向預計出貨明細頁面，或回傳編輯頁面與錯誤訊息。</returns>
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Edit_SC(int sno, string sales_order, string cust_wono, int? exp_shipquantity
            , string sap_mark,int? stock_quantity, string sap_in, string company_code, string wono
            , string position, string eng_sr, int ssno, int box_quantity, string transportation)
        {
            // 權限不足或未登入
            var authResult = Login_Authentication();
            if (authResult != null)
            {
                return authResult;
            }

            //取得會員帳號指定fUserId
            string UserId = ViewBag.UserId;

            //清除公司代碼
            if (company_code == "")
            {
                #region 預計出貨消除公司代碼
                //不再在編輯時扣帳因此取消數量異動邏輯
                //var tstock = db.E_StoreHouseStock.Where(w => w.wono == wono && w.sno == sno).FirstOrDefault();
                //tstock.quantity += exp_shipquantity;
                //db.SaveChanges();

                var _sc = db.E_StoreHouseStock_SC.Find(ssno);
                _sc.company_name = "";
                _sc.address = "";
                _sc.company_code = "";
                db.SaveChanges();
                #endregion
            }
            else
            {
                //檢查數量的邏輯
                if ((int)(stock_quantity - exp_shipquantity) < 0)
                {
                    Edit_View(sno, sales_order);
                    return View("Edit_SC", "_LayoutMember", Edit_View(sno, sales_order));
                }
                else
                {
                    var comp = db.E_Compyany.Where(s => s.code == company_code).ToList();
                    string sdate = (from E_StoreHouseStock in db.E_StoreHouseStock
                                    where E_StoreHouseStock.sno == sno
                                    select E_StoreHouseStock.inputdate).FirstOrDefault();

                    //db.E_StoreHouseStock.Where(s => s.sno == sno).FirstOrDefault();
                    //找出未成為訂單明細資料,即購物車內容(m.wono).Substring(0, 7)
                    //var stock_Order = db.E_StoreHouseStock_SC
                    //    .Where(s => s.sales_order == wono.Substring(0, 7) && s.eng_sr== eng_sr).FirstOrDefault();
                    //找不到公司資料就不完成編輯
                    if (comp.Count > 0)
                    {
                        var stock_Order = db.E_StoreHouseStock_SC
                        .Where(s => s.sno == ssno).FirstOrDefault();

                        foreach (var item in comp)
                        {
                            stock_Order.wono = wono;
                            stock_Order.cust_wono = cust_wono;
                            stock_Order.stock_quantity = stock_quantity;
                            stock_Order.box_quantity = box_quantity;
                            stock_Order.exp_shipquantity = exp_shipquantity;
                            stock_Order.sap_mark = sap_mark;
                            stock_Order.sap_in = sap_in;
                            stock_Order.company_name = item.company_Name;
                            stock_Order.address = item.address;
                            stock_Order.company_code = company_code;
                            // 20250821 By Jesse 修改 stock_Order.inputdate = sdate; 顯示 yyyy-MM-dd 格式
                            stock_Order.inputdate = DateTime.TryParse(sdate, out DateTime dt) ? dt.ToString("yyyy-MM-dd") : sdate;
                            stock_Order.position = position;
                            stock_Order.transportation = transportation;
                        }

                        try
                        {
                            db.SaveChanges();
                        }
                        catch (Exception)
                        {
                            // 保持原本行為：向上拋出
                            throw;
                        }

                        #region 依據wono找出庫存表的工單進行扣帳;表E_StoreHouseStock
                        //var tstock = db.E_StoreHouseStock.Where(w => w.sno == sno && w.position == position).FirstOrDefault();
                        //tstock.quantity = tstock.quantity - exp_shipquantity;
                        //db.SaveChanges();
                        #endregion
                    }
                    else
                    {
                        ViewBag.err = "找不到公司代號!";
                        return View("Edit_SC", "_LayoutMember", Edit_View(sno, sales_order));
                    }
                }
            }
            return RedirectToAction("StoreHouseStock_SC");

        }

        /// <summary>
        /// 確認出貨，將[預計出貨]明細資料正式轉為訂單，並處理庫存扣帳與已銷單庫存。
        /// <para>1. 取得[預計出貨]-公司名稱不為空及未確認。</para>
        /// <para>2. 取得會員帳號，建立訂單識別碼。</para>
        /// <para>3. 將購物車狀態的預計出貨明細（IsApproved="N"）設為已確認（IsApproved="Y"），並加入訂單識別碼。</para>
        /// <para>4. 建立出貨明細（E_StoreHouseStock_Order），並將公司代號為5的資料寫入已銷單庫存（E_StoreHouseStock_BOS）。</para>
        /// <para>5. 確認出貨後，依據出貨明細扣除成倉庫存數量。</para>
        /// </summary>
        /// <param name="sno">訂單唯一識別碼（未使用，保留參數）</param>
        /// <returns>出貨明細頁面</returns>
        [HttpPost]
        public ActionResult OrderConfirm(string sno)
        {
            // 權限不足或未登入
            var authResult = Login_Authentication();
            if (authResult != null)
            {
                return authResult;
            }

            // 找出會員帳號指定給fUserId
            string fUserId = ViewBag.UserId;
            //建立唯一的識別,並指定給guid變數用來當作訂單編號
            //tOrder的fOrderGuid欄位會關連到tOrderDetail的fOrderGuid欄位
            //形成一對的關係,即一筆訂單資料對應多筆訂單明細(Master-Detail)
            string guid = Guid.NewGuid().ToString();

            //找出目前在預計出貨中"未有識別碼"及"公司"不為空的資料
            var carList = db.E_StoreHouseStock_SC.Where(m => m.IsApproved == "N" && m.company_name != "").ToList();
            //將購物車狀態的fIsApproved設定為"是"表示確認訂購產品
            foreach (var item in carList)
            {
                if (!string.IsNullOrEmpty(item.company_name) || item.company_code == "5")
                {
                    item.OrderGuid = guid;     //進入該迴圈表示訂單成立,加入訂單識別碼
                    item.IsApproved = "Y";
                    db.Entry(item).State = EntityState.Modified;
                }
            }
            //更新資料庫, 異動tOrder 和 tOrderDetail
            //完成訂單主檔和訂單明細的更新
            try
            {
                db.SaveChanges();
            }
            catch (Exception)
            {
                // 保持原本行為：向上拋出
                throw;
            }
            #region 預計出貨資料 - E_StoreHouseStock_SC
            //找出當日預計出貨資料(Y+當日)
            string today = DateTime.Now.ToString("yyyy-MM-dd");
            //寫出貨明細(加入訂單碼)
            var Y_carList = db.E_StoreHouseStock_SC
                                .Where(m => m.IsApproved == "Y" //&& m.company_code != "5"
                                && m.exp_shipdate == today && !string.IsNullOrEmpty(m.OrderGuid)
                                && !db.E_StoreHouseStock_Order.Any(x => x.OrderGuid == m.OrderGuid)).ToList();
            //寫已銷單庫存(不加入訂單碼)
            var carList_5 = db.E_StoreHouseStock_SC
                                .Where(m => m.IsApproved == "Y" && m.company_code == "5"
                                && m.exp_shipdate == today && !string.IsNullOrEmpty(m.OrderGuid)
                                && !db.E_StoreHouseStock_Order.Any(x => x.OrderGuid == m.OrderGuid)).ToList();
            #endregion
            //建立當日訂單主檔資料 company_code != "5"
            foreach (var Y_item in Y_carList)
            {
                E_StoreHouseStock_Order order = new E_StoreHouseStock_Order();
                order.OrderGuid = Y_item.OrderGuid;    //訂單識別碼
                order.UserId = fUserId;
                order.wono = Y_item.wono.Trim();
                order.cust_wono = Y_item.cust_wono.Trim();
                order.eng_sr = Y_item.eng_sr.Trim();
                order.sales_order = Y_item.sales_order.Trim();
                order.box_quantity = Y_item.box_quantity;
                order.exp_shipquantity = Y_item.exp_shipquantity;
                order.stock_quantity = Y_item.stock_quantity;
                order.sap_mark = Y_item.sap_mark.Trim();
                order.sap_in = Y_item.sap_in;
                order.company_code = Y_item.company_code.Trim();
                order.company_name = Y_item.company_name;
                order.address = Y_item.address;
                order.transportation = Y_item.transportation;
                order.chk_date = DateTime.Now;
                order.position = Y_item.position;
                try
                {
                    db.E_StoreHouseStock_Order.Add(order);
                }
                catch (Exception)
                {
                    // 保持原本行為：向上拋出
                    throw;
                }
            }
            db.SaveChanges();

            #region 寫入已銷單庫存 - 預計出貨資料 - 公司代號 = 5 (已銷單庫存)
            foreach (var item_5 in carList_5)
            {
                //檢查E_StoreHouseStock_BOS是否已有該機種資料，考量到會異動庫位
                //20251113 Jesse 修改多加庫位條件
                var engsr_bos = db.E_StoreHouseStock_BOS
                    .FirstOrDefault(x => x.eng_sr == item_5.eng_sr && x.position == item_5.position);
                if (engsr_bos != null)
                {
                    engsr_bos.quantity += item_5.exp_shipquantity;
                }
                else
                {
                    E_StoreHouseStock_BOS stock_BOS = new E_StoreHouseStock_BOS();

                    stock_BOS.quantity = item_5.exp_shipquantity;
                    stock_BOS.position = item_5.position;
                    stock_BOS.eng_sr = item_5.eng_sr.Trim();
                    stock_BOS.ship_quantity = 0;
                    stock_BOS.transportation = item_5.transportation.Trim();
                    stock_BOS.date = item_5.exp_shipdate;
                    try
                    {
                        db.E_StoreHouseStock_BOS.Add(stock_BOS);
                    }
                    catch (Exception)
                    {
                        // 保持原本行為：向上拋出
                        throw;
                    }
                }
            }
            db.SaveChanges();
            #endregion
            #region XX寫入已銷單庫存XX
            //找出公司代號5的資料
            //var result = (from a in db.E_StoreHouseStock_SC
            //              join b in db.E_StoreHouseStock on a.eng_sr equals b.eng_sr into b_join
            //              from b in b_join.DefaultIfEmpty()
            //              join c in db.E_Weight on new { Eng_sr = a.eng_sr } equals new { Eng_sr = c.EngSr } into c_join
            //              from c in c_join.DefaultIfEmpty()
            //              where
            //                a.IsApproved == "Y" &&
            //                a.company_code == "5" &&
            //                a.exp_shipdate == today
            //              select new
            //              {
            //                  a.eng_sr,
            //                  a.exp_shipdate,
            //                  a.exp_shipquantity,
            //                  Position = b.position,
            //                  a.transportation,
            //                  Full_Amount = (int?)c.Full_Amount,
            //                  Full_GW = (decimal?)c.Full_GW,
            //                  Full_NW = (decimal?)c.Full_NW,
            //              }).Distinct();
            //if (result.Count() > 0)
            //{
            //    List<E_StoreHouseStock_BOS> stock_BOS_List = new List<E_StoreHouseStock_BOS>();
            //    //var groupedResult = result.GroupBy(x => x.eng_sr);

            //    foreach (var group in result)
            //    {
            //        var engsr_bos = db.E_StoreHouseStock_BOS.FirstOrDefault(x => x.eng_sr == group.eng_sr);
            //        if (engsr_bos == null)
            //        {
            //            E_StoreHouseStock_BOS stock_BOS = new E_StoreHouseStock_BOS();
            //            stock_BOS.eng_sr = group.eng_sr;
            //            stock_BOS.date = group.exp_shipdate;
            //            stock_BOS.position = group.Position;
            //            stock_BOS.quantity = group.exp_shipquantity;
            //            stock_BOS.Full_Amount = group.Full_Amount;
            //            stock_BOS.Full_GW = group.Full_GW;
            //            stock_BOS.Full_NW = group.Full_NW;
            //            stock_BOS.ship_quantity = 0;
            //            stock_BOS.transportation = group.transportation;
            //            stock_BOS_List.Add(stock_BOS);
            //        }
            //        else
            //        {
            //            engsr_bos.quantity += group.exp_shipquantity;
            //        }
            //    }
            //    try
            //    {
            //        db.E_StoreHouseStock_BOS.AddRange(stock_BOS_List);
            //        db.SaveChanges();
            //    }
            //    catch (Exception torder)
            //    {
            //        throw;
            //    }
            //}
            #endregion

            #region 取消編輯後扣帳改為確認出貨再扣帳
            //Y_carList 代表編輯後這一批要出貨的帳
            foreach (var item in Y_carList)
            {
                if (item.wono != "無工單出貨")
                {
                    var tstock = db.E_StoreHouseStock
                    .Where(w => w.wono == item.wono && w.position == item.position && w.del_flag != "D").FirstOrDefault();
                    tstock.quantity = tstock.quantity - item.exp_shipquantity;
                    db.SaveChanges();
                }

            }
            //carList_5 代表公司代碼5的扣帳
            foreach (var item_5 in carList_5)
            {
                if (item_5.wono != "無工單出貨")
                {
                    var tstock_5 = db.E_StoreHouseStock
                    .Where(w => w.eng_sr == item_5.eng_sr && w.position == item_5.position && item_5.wono == "").FirstOrDefault();
                    if (tstock_5 != null)
                    {
                        tstock_5.quantity = tstock_5.quantity - item_5.exp_shipquantity;
                        db.SaveChanges();
                    }
                }

            }

            #endregion

            return RedirectToAction("StoreHouseStock_Order");
        }
        #endregion

        #region 無單號出貨
        /// <summary>
        /// Edits the s cbos.
        /// </summary>
        /// <param name="sno">The sno.</param>
        /// <param name="sales_order">The sales order.</param>
        /// <returns></returns>
        public ActionResult Edit_SCbos(int sno, string sales_order)
        {
            // 權限不足或未登入
            var authResult = Login_Authentication();
            if (authResult != null)
            {
                return authResult;
            }

            //取得會員帳號指定fUserId
            string UserId = ViewBag.UserId;

            var bos_sc = db.E_StoreHouseStock_SC.Where(b => b.sno == sno).FirstOrDefault();

            return View("Edit_SCbos", "_LayoutMember", bos_sc);
        }

        /// <summary>
        /// 編輯已銷單預計出貨明細（BOS）。
        /// <para>1. 若 company_code 為空，則將公司相關欄位清空，並將已銷單庫存數量加回。</para>
        /// <para>2. 若 company_code 不為空，則更新出貨明細的公司資訊與運輸方式。</para>
        /// <para>3. 已銷單庫存的扣帳已在其他流程處理。</para>
        /// </summary>
        /// <param name="sno">預計出貨明細的唯一識別碼。</param>
        /// <param name="exp_shipquantity">預計出貨數量。</param>
        /// <param name="sap_mark">SAP 備註。</param>
        /// <param name="stock_quantity">成倉庫存數量。</param>
        /// <param name="company_code">公司代號。</param>
        /// <param name="position">庫位。</param>
        /// <param name="eng_sr">機種名稱。</param>
        /// <param name="transportation">運輸方式。</param>
        /// <returns>編輯完成後導向預計出貨明細頁面。</returns>
        [HttpPost]
        public ActionResult Edit_SCbos(int sno, int? exp_shipquantity, string sap_mark,
            int? stock_quantity, string company_code, string position, string eng_sr, string transportation)
        {
            // 權限不足或未登入
            var authResult = Login_Authentication();
            if (authResult != null)
            {
                return authResult;
            }

            //取得會員帳號指定fUserId
            string UserId = ViewBag.UserId;

            if (company_code == "")
            {
                #region 預計出貨消除公司代碼
                var tstock = db.E_StoreHouseStock_BOS.Where(w => w.eng_sr == eng_sr).FirstOrDefault();
                tstock.quantity += exp_shipquantity;
                db.SaveChanges();

                var _sc = db.E_StoreHouseStock_SC.Find(sno);
                _sc.company_name = "";
                _sc.address = "";
                _sc.company_code = "";
                db.SaveChanges();
                #endregion
            }
            else
            {
                var comp = db.E_Compyany.Where(s => s.code == company_code).ToList();

                //db.E_StoreHouseStock.Where(s => s.sno == sno).FirstOrDefault();
                //找出未成為訂單明細資料,即購物車內容(m.wono).Substring(0, 7)

                var stock_Order = db.E_StoreHouseStock_SC
                    .Where(s => s.sno == sno).FirstOrDefault();

                foreach (var item in comp)
                {
                    stock_Order.transportation = transportation;
                    stock_Order.sap_mark = "";
                    stock_Order.company_name = item.company_Name;
                    stock_Order.address = item.address;
                    stock_Order.company_code = company_code;
                }

                try
                {
                    db.SaveChanges();
                }
                catch (Exception)
                {
                    // 保持原本行為：向上拋出
                    throw;
                }

                //已經在表E_StoreHouseStock_BOS扣完帳
            }
            return RedirectToAction("StoreHouseStock_SC");

        }
        #endregion

        //AJAX 更新庫位資訊
        /// <summary>
        /// 依據機種名稱 (eng_sr) 取得成倉庫存資料，回傳 JSON 格式。
        /// </summary>
        /// <param name="eng_sr">機種名稱</param>
        /// <returns>
        /// 若找到資料則回傳 JSON 格式的庫存資料，否則回傳 404 Not Found。
        /// </returns>
        public ActionResult Edit_UpdateBOS(string eng_sr)
        {
            //取得會員帳號指定fUserId
            //string UserId = (Session["Member"] as MaintainViewModels).fUserId;
            var stocks = db.E_StoreHouseStock.Where(m => m.eng_sr == eng_sr && m.del_flag != "D").FirstOrDefault();
            if (stocks == null)
            {
                return HttpNotFound();
            }
            return Json(stocks, JsonRequestBehavior.AllowGet);
        }

        /// <summary>
        /// 依據工單序號 (sno) 查詢成倉庫存資料，回傳 JSON 格式。
        /// </summary>
        /// <param name="sno">工單序號</param>
        /// <returns>
        /// 若找到資料則回傳 JSON 格式的庫存資料，否則回傳 404 Not Found。
        /// </returns>
        public ActionResult Edit_Update(int sno)
        {
            //取得會員帳號指定fUserId
            //string UserId = (Session["Member"] as MaintainViewModels).fUserId;
            var stocks = db.E_StoreHouseStock.Where(m => m.sno == sno && m.del_flag != "D").FirstOrDefault();
            if (stocks == null)
            {
                return HttpNotFound();
            }
            return Json(stocks, JsonRequestBehavior.AllowGet);
        }

        #region 出貨明細
        /// <summary>
        /// 出貨明細
        /// </summary>
        /// <returns></returns>
        public ActionResult StoreHouseStock_Order()
        {
            // 權限不足或未登入
            var authResult = Login_Authentication();
            if (authResult != null)
            {
                return authResult;
            }

            // 找出當日預計出貨資料
            string today = DateTime.Now.ToString("yyyy-MM-dd");
            var result = db.E_StoreHouseStock_Order
                .Where(o => o.chk_date.ToString() == today && o.del_flag != "D").OrderBy(o => o.sap_mark).ToList();
            return View("StoreHouseStock_Order", "_LayoutMember", result);
        }
        #region 匯出當日出貨明細為 Excel 檔案
        /// <summary>
        /// 匯出當日出貨明細為 Excel 檔案。
        /// </summary>
        /// <returns>
        /// 下載 Excel 檔案（出貨明細.xlsx），內容包含當日所有出貨明細資料。
        /// </returns>
        public ActionResult ExportToExcel()
        {
            string today = DateTime.Now.ToString("yyyy-MM-dd");
            // 取得出貨明細資料
            var orders = db.E_StoreHouseStock_Order.Where(o => o.chk_date.ToString() == today).OrderBy(o => o.sap_mark).ToList();

            // 建立 Excel 工作簿
            XSSFWorkbook workbook = new XSSFWorkbook();
            ISheet sheet = workbook.CreateSheet("出貨明細");

            // 建立 Excel 表頭
            IRow headerRow = sheet.CreateRow(0);
            headerRow.CreateCell(0).SetCellValue("工單");
            headerRow.CreateCell(1).SetCellValue("機種");
            headerRow.CreateCell(2).SetCellValue("業務單號");
            headerRow.CreateCell(3).SetCellValue("客戶單號");
            headerRow.CreateCell(4).SetCellValue("預計出貨數");
            headerRow.CreateCell(5).SetCellValue("SAP備註");
            headerRow.CreateCell(6).SetCellValue("出貨方式");
            headerRow.CreateCell(7).SetCellValue("成倉數量");
            //headerRow.CreateCell(7).SetCellValue("SAP入庫");
            headerRow.CreateCell(8).SetCellValue("公司名稱");
            headerRow.CreateCell(9).SetCellValue("地址");

            // 填入 Excel 資料
            int rowIndex = 1;
            foreach (var order in orders)
            {
                IRow dataRow = sheet.CreateRow(rowIndex);
                dataRow.CreateCell(0).SetCellValue(order.wono);
                dataRow.CreateCell(1).SetCellValue(order.eng_sr);
                dataRow.CreateCell(2).SetCellValue(order.sales_order);
                dataRow.CreateCell(3).SetCellValue(order.cust_wono);
                dataRow.CreateCell(4).SetCellValue(order.exp_shipquantity.ToString());
                dataRow.CreateCell(5).SetCellValue(order.sap_mark);
                dataRow.CreateCell(6).SetCellValue(order.transportation);
                dataRow.CreateCell(7).SetCellValue(order.stock_quantity.ToString());
                //dataRow.CreateCell(7).SetCellValue(order.sap_in.ToString());
                dataRow.CreateCell(8).SetCellValue(order.company_name);
                dataRow.CreateCell(9).SetCellValue(order.address);
                rowIndex++;
            }

            // 設定檔案名稱
            string fileName = "出貨明細.xlsx";

            // 儲存 Excel 檔案
            using (MemoryStream stream = new MemoryStream())
            {
                workbook.Write(stream);
                var content = stream.ToArray();
                return File(content, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
            }
        }
        #endregion
        #endregion

        /// <summary>
        /// 刪除出貨訂單，並將庫存數量加回，記錄刪除原因。
        /// </summary>
        /// <param name="sno">出貨訂單的唯一識別碼。</param>
        /// <param name="reason">刪除原因。</param>
        /// <param name="wono">工單號碼。</param>
        /// <param name="quantity">需加回的庫存數量。</param>
        /// <param name="guid">訂單識別碼。</param>
        /// <returns>導向庫存查詢頁面。</returns>
        public ActionResult DeleteOrder(int sno, string reason, string wono, int quantity, string guid)
        {
            var record = db.E_StoreHouseStock_Order.Find(sno);
            if (record != null)
            {
                record.del_flag = "D";
                record.record = reason;
                //db.E_StoreHouseStock_Order.Add(record);
                db.SaveChanges();
            }
            #region 從庫存扣的帳要加回去
            /*
              var Y_carList = db.E_StoreHouseStock_SC
                                        .Where(m => m.IsApproved == "Y" //&& m.company_code != "5"
                                        && m.exp_shipdate == today && !string.IsNullOrEmpty(m.OrderGuid)
                                        && !db.E_StoreHouseStock_Order.Any(x => x.OrderGuid == m.OrderGuid)).ToList();
             */

            var tstock = db.E_StoreHouseStock.Where(w => w.wono == wono && w.position == record.position).FirstOrDefault();
            tstock.quantity += quantity;
            db.SaveChanges();
            #endregion

            #region 退回去預計出貨
            //var scstock = db.E_StoreHouseStock_SC
            //    .Where(sc => sc.OrderGuid == guid && sc.wono==wono).FirstOrDefault();
            //scstock.company_name = "";
            //scstock.address = "";
            //scstock.IsApproved = "N";
            //db.SaveChanges();
            #endregion

            return RedirectToAction("StoreHouseStock");
        }

        //bill of sale 銷貨單據,銷單
        #region 已銷單庫存
        /// <summary>
        /// 已銷單庫存
        /// </summary>
        /// <param name="engsr">機種名稱查詢條件</param>
        /// <returns></returns>
        public ActionResult StoreHouseStock_BOS(string engsr = null)
        {
            // 權限不足或未登入
            var authResult = Login_Authentication();
            if (authResult != null)
            {
                return authResult;
            }

            if (!string.IsNullOrEmpty(Session["Member"].ToString()))
            {
                string sql = @"SELECT Distinct eng_sr,quantity,position,transportation,a.Full_Amount, a.Full_GW, a.Full_NW, b.sno ,a.sno 
                                FROM E_StoreHouseStock_BOS a LEFT JOIN E_Weight b on a.eng_sr=b.EngSr WHERE a.quantity > 0 ";
                List<SqlParameter> paramList = new List<SqlParameter>();
                if (!string.IsNullOrWhiteSpace(engsr))
                {
                    sql += " AND a.eng_sr LIKE @engsr ";
                    paramList.Add(new SqlParameter("@engsr", engsr + "%"));
                }
                DataSet dataSet = dbMethod.ExecuteDataSet(sql, CommandType.Text, paramList.Count > 0 ? paramList.ToArray() : null);
                return View("StoreHouseStock_BOS", "_LayoutMember", dataSet);
            }
            return View();
        }

        #region 匯出已銷單庫存 Excel
        /// <summary>
        /// 匯出已銷單庫存 Excel
        /// </summary>
        /// <param name="engsr">機種名稱查詢條件</param>
        public ActionResult ExportBOSExcel(string engsr)
        {
            string sql = @"SELECT Distinct eng_sr,quantity,position, b.sno ,a.sno
                            FROM E_StoreHouseStock_BOS a
                            LEFT JOIN E_Weight b ON a.eng_sr=b.EngSr WHERE a.quantity > 0 ";
            List<SqlParameter> paramList = new List<SqlParameter>();
            if (!string.IsNullOrWhiteSpace(engsr))
            {
                sql += " AND a.eng_sr LIKE @engsr ";
                paramList.Add(new SqlParameter("@engsr", engsr + "%"));
            }
            DataTable dt = dbMethod.ExecuteDataTable(sql, CommandType.Text, paramList.ToArray());

            if (dt == null || dt.Rows.Count == 0)
            {
                // 沒有資料時回傳提示訊息或空檔案
                TempData["ErrorMessage"] = "查無資料，無法匯出 Excel。";
                return RedirectToAction("StoreHouseStock_BOS");
            }

            XSSFWorkbook workbook = new XSSFWorkbook();
            try
            {
                var sheet = workbook.CreateSheet("已銷單庫存");
                var headerRow = sheet.CreateRow(0);
                string[] headers = { "機種名稱", "數量", "庫位" };
                for (int i = 0; i < headers.Length; i++)
                {
                    headerRow.CreateCell(i).SetCellValue(headers[i]);
                }
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    var row = sheet.CreateRow(i + 1);
                    row.CreateCell(0).SetCellValue(dt.Rows[i]["eng_sr"].ToString());
                    row.CreateCell(1).SetCellValue(dt.Rows[i]["quantity"].ToString());
                    row.CreateCell(2).SetCellValue(dt.Rows[i]["position"].ToString());
                }
                using (var exportData = new MemoryStream())
                {
                    workbook.Write(exportData);
                    string fileName = $"已銷單庫存_{DateTime.Now:yyyyMMddHHmmss}.xlsx";
                    return File(exportData.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
                }
            }
            finally
            {
                workbook.Close();
            }
        }
        #endregion

        #region 已銷出貨
        /// <summary>
        /// 已銷出貨 - 進入已銷單庫存出貨頁面
        /// </summary>
        /// <param name="eng_sr">機種名稱</param>  // 指定要查詢的機種名稱
        /// <param name="position">庫位</param>    // 指定要查詢的庫位
        /// <returns>回傳已銷單庫存出貨頁面</returns> // 回傳對應的 View
        public ActionResult BOS_toShip(string eng_sr, string position)
        {
            var _bos = db.E_StoreHouseStock_BOS.Where(m => m.eng_sr == eng_sr && m.position == position).FirstOrDefault();
            return View("BOS_toShip", "_LayoutMember", _bos);
        }
        /// <summary>
        /// 已銷出貨 - 已銷單庫存出貨作業：
        /// 1. 依據機種名稱與庫位，扣除已銷單庫存數量，並記錄出貨數量與出貨日期。
        /// 2. 將出貨資料寫入 E_StoreHouseStock_SC（預計出貨明細），以利後續出貨流程。
        /// </summary>
        /// <param name="eng_sr">機種名稱</param>
        /// <param name="ship_quantity">出貨數量</param>
        /// <param name="position">庫位</param>
        /// <param name="transportation">運輸方式</param>
        /// <returns>出貨完成後導向已銷單庫存頁面，若失敗則回傳原畫面並顯示錯誤</returns>
        [HttpPost]
        public ActionResult BOS_toShip(string eng_sr, int ship_quantity, string position, string transportation)
        {
            // 權限不足或未登入
            var authResult = Login_Authentication();
            if (authResult != null)
            {
                return authResult;
            }

            // 取得會員帳號指定 fUserId
            string UserId = ViewBag.UserId;

            using (var transaction = db.Database.BeginTransaction())
            {
                try
                {
                    // 已銷單庫存-找出符合條件的資料 - 機種名稱與庫位
                    var bos = db.E_StoreHouseStock_BOS
                        .Where(m => m.eng_sr == eng_sr && m.position == position)
                        .ToList();

                    if (bos.Count > 0)
                    {
                        string today = DateTime.Now.ToString("yyyy-MM-dd");

                        // 更新-已銷單庫存- E_StoreHouseStock_BOS 資料
                        foreach (var item in bos)
                        {
                            item.ship_quantity = ship_quantity;
                            item.quantity -= ship_quantity;
                            item.ship_date = today;
                            item.transportation = transportation;
                        }

                        // 儲存異動
                        db.SaveChanges();

                        // 將資料-已銷單庫存-新增到-預計出貨- E_StoreHouseStock_SC 中
                        var tos = bos.Select(item => new E_StoreHouseStock_SC
                        {
                            exp_shipdate = item.ship_date,
                            eng_sr = item.eng_sr,
                            exp_shipquantity = item.ship_quantity,
                            stock_quantity = item.quantity + item.ship_quantity,
                            position = item.position,
                            transportation = item.transportation,
                            IsApproved = "N",
                            sales_order = "N/A",
                            cust_wono = "N/A",
                            wono = "無工單出貨"
                        }).ToList();

                        // 儲存異動
                        db.E_StoreHouseStock_SC.AddRange(tos);
                        db.SaveChanges();

                        // 提交事務
                        transaction.Commit();
                    }

                    return RedirectToAction("StoreHouseStock_BOS");
                }
                catch (Exception ex)
                {
                    // 回復
                    transaction.Rollback();
                    // 處理錯誤
                    ModelState.AddModelError("", ex.Message);
                    return View();
                }
            }
        }
        #endregion
        #region 庫位異動
        /// <summary>
        /// 編輯已銷單庫存資料的 View。
        /// </summary>
        /// <param name="eng_sr">機種名稱</param>
        /// <returns>回傳 Edit_BOS 頁面，並帶入指定機種的已銷單庫存資料</returns>
        public ActionResult Edit_BOS(string eng_sr)
        {
            // 權限不足或未登入
            var authResult = Login_Authentication();
            if (authResult != null)
            {
                return authResult;
            }

            //取得會員帳號指定fUserId
            string UserId = ViewBag.UserId;

            var _bos = db.E_StoreHouseStock_BOS.Where(m => m.eng_sr == eng_sr && m.quantity > 0).FirstOrDefault();
            return View("Edit_BOS", "_LayoutMember", _bos);
        }

        /// <summary>
        /// 編輯已銷單庫存資料-庫位異動。
        /// 1. 若目標庫位 (changeP) 已存在相同機種名稱的資料，則將數量加總。
        /// 2. 若目標庫位不存在，則將原始庫位資料複製並把庫位變更為目標庫位，新增一筆資料。
        /// 3. 原始庫位資料數量歸零。
        /// </summary>
        /// <param name="eng_sr">機種名稱</param>
        /// <param name="position">原始庫位</param>
        /// <param name="changeP">目標庫位</param>
        /// <param name="quantity">要轉移的數量</param>
        /// <returns>轉移完成後導向已銷單庫存頁面</returns>
        [HttpPost]
        public ActionResult Edit_BOS(string eng_sr, string position, string changeP, int quantity)
        {
            // 權限驗證
            var authResult = Login_Authentication();
            if (authResult != null)
            {
                return authResult;
            }

            // 參數防呆檢查，錯誤訊息寫入 ModelState 以便前端顯示
            if (string.IsNullOrWhiteSpace(eng_sr) || string.IsNullOrWhiteSpace(position) || string.IsNullOrWhiteSpace(changeP))
            {
                ModelState.AddModelError("", "參數不可為空，請重新操作。");
                var _bos = db.E_StoreHouseStock_BOS.Where(m => m.eng_sr == eng_sr && m.quantity > 0).FirstOrDefault();
                return View("Edit_BOS", "_LayoutMember", _bos);
            }
            if (quantity <= 0)
            {
                ModelState.AddModelError("", "轉移數量必須大於0。");
                var _bos = db.E_StoreHouseStock_BOS.Where(m => m.eng_sr == eng_sr && m.quantity > 0).FirstOrDefault();
                return View("Edit_BOS", "_LayoutMember", _bos);
            }
            if (position == changeP)
            {
                ModelState.AddModelError("", "原始庫位與目標庫位不可相同。");
                var _bos = db.E_StoreHouseStock_BOS.Where(m => m.eng_sr == eng_sr && m.quantity > 0).FirstOrDefault();
                return View("Edit_BOS", "_LayoutMember", _bos);
            }

            // 取得會員帳號指定fUserId
            string UserId = ViewBag.UserId;

            using (var db = new E_StoreHouseEntities())
            using (var transaction = db.Database.BeginTransaction())
            {
                try
                {
                    // 取得原始庫位資料且數量足夠
                    var srcBos = db.E_StoreHouseStock_BOS.FirstOrDefault(m => m.eng_sr == eng_sr && m.position == position && m.quantity > 0);
                    if (srcBos == null)
                    {
                        ModelState.AddModelError("", "找不到原始庫位資料或數量不足。");
                        var _bos = db.E_StoreHouseStock_BOS.Where(m => m.eng_sr == eng_sr && m.quantity > 0).FirstOrDefault();
                        return View("Edit_BOS", "_LayoutMember", _bos);
                    }
                    if (srcBos.quantity < quantity)
                    {
                        ModelState.AddModelError("", "原始庫位數量不足，無法轉移。");
                        var _bos = db.E_StoreHouseStock_BOS.Where(m => m.eng_sr == eng_sr && m.quantity > 0).FirstOrDefault();
                        return View("Edit_BOS", "_LayoutMember", _bos);
                    }

                    // 檢查目標庫位是否已有相同機種資料
                    var destBos = db.E_StoreHouseStock_BOS.SingleOrDefault(m => m.eng_sr == eng_sr && m.position == changeP);
                    if (destBos != null)
                    {
                        // 目標庫位已有資料，直接加總數量
                        destBos.quantity += quantity;
                    }
                    else
                    {
                        // 目標庫位無資料，複製原始資料並設置新庫位與數量
                        var newBos = new E_StoreHouseStock_BOS
                        {
                            eng_sr = srcBos.eng_sr,
                            position = changeP,
                            date = srcBos.date,
                            quantity = quantity,
                            Full_Amount = srcBos.Full_Amount,
                            Full_GW = srcBos.Full_GW,
                            Full_NW = srcBos.Full_NW,
                            ship_quantity = srcBos.ship_quantity,
                            ship_date = srcBos.ship_date,
                            transportation = srcBos.transportation
                        };
                        db.E_StoreHouseStock_BOS.Add(newBos);
                    }

                    // 原始庫位數量減少，若歸零則設為0
                    srcBos.quantity -= quantity;
                    if (srcBos.quantity < 0) srcBos.quantity = 0;

                    db.SaveChanges();
                    transaction.Commit();
                }
                catch (Exception ex)
                {
                    transaction.Rollback();
                    // 記錄錯誤訊息於 ModelState 以便前端顯示
                    ModelState.AddModelError("", "異動失敗：" + ex.Message);
                    var _bos = db.E_StoreHouseStock_BOS.Where(m => m.eng_sr == eng_sr && m.quantity > 0).FirstOrDefault();
                    return View("Edit_BOS", "_LayoutMember", _bos);
                }
            }
            return RedirectToAction("StoreHouseStock_BOS");
        }
        #endregion

        #endregion

        #region 資料查詢
        /// <summary>
        /// 查詢出貨明細與 QC 資料。
        /// </summary>
        /// <param name="indate">起始日期</param>
        /// <param name="indate2">結束日期</param>
        /// <param name="engsr">機種名稱</param>
        /// <param name="wono">工單號碼</param>
        /// <param name="mark">備註</param>
        /// <param name="li">清單類型</param>
        /// <param name="order_cust">客戶業單</param>
        /// <param name="wono_cust">客戶工單</param>
        /// <returns>回傳出貨明細與 QC 資料的頁面</returns>
        public ActionResult StoreHouseStock_OrderData(string indate, string indate2, string engsr
            , string wono, string mark, string li, string order_cust, string wono_cust)
        {
            // 權限不足或未登入
            var authResult = Login_Authentication();
            if (authResult != null)
            {
                return authResult;
            }
            //取得StockData
            OrderData_Model oDM = new OrderData_Model();
            oDM.indate = indate;
            oDM.indate2 = indate2;
            oDM.engsr = engsr;
            oDM.wono = wono;
            oDM.mark = mark;
            oDM.li = li;
            oDM.order_cust = order_cust;
            oDM.wono_cust = wono_cust;
            DataSet StockData = oDM.Gds();

            //取得QCData
            //QC_Model qc = new QC_Model(db);
            //List<QC_Model.QCData> QCData = qc.GetData();
            //int SumProCount = QCData.Sum(s => s.ProCount);


            //建立ViewModel
            var viewModel = new StockOrderDataAndQcDataViewModels
            {
                StockData = StockData,
                //QCData = qc.GetData(),
            };

            //int SumProCount = viewModel.QCData.Sum(s => s.ProCount);
            //ViewBag.Total = SumProCount;
            ViewBag.List = li;
            return View("StoreHouseStock_OrderData", "_LayoutMember", viewModel);
        }
        #endregion

        #region 預計直出資料

        // 新增預計直出資料查詢條件：wono, eng_sr, startDate, endDate By 20251007 Jesse
        /// <summary>
        /// 預計直出查詢（支援依 wono / eng_sr / startDate / endDate 篩選）
        /// 優化重點：
        /// 1. 將 SQL 建構邏輯抽整合以減少重複字串。
        /// 2. 先回傳 checkOK=true 的資料（優先顯示），再接著回傳 checkOK=false 的資料，最後以複合鍵去重。
        /// 3. 使用參數化查詢避免 SQL 注入並提升可維護性。
        /// </summary>
        /// <param name="wono">工單號碼（模糊查詢）</param>
        /// <param name="eng_sr">機種名稱（模糊查詢）</param>
        /// <param name="startDate">起始日期（yyyy-MM-dd）</param>
        /// <param name="endDate">結束日期（yyyy-MM-dd）</param>
        /// <returns>回傳 StoreHouseDropshipping View，模型為 EDropshippingQDto 列表</returns>
        public ActionResult StoreHouseDropshipping(string wono, string eng_sr, string startDate, string endDate)
        {
            // 權限檢查
            var authResult = Login_Authentication();
            if (authResult != null)
            {
                return authResult;
            }

            List<EDropshippingQDto> finalResult = new List<EDropshippingQDto>();

            try
            {
                finalResult = GetDropshipResults(wono, eng_sr, startDate, endDate);
            }
            catch (Exception ex)
            {
                ModelState.AddModelError("", ex.Message);
                finalResult = new List<EDropshippingQDto>();
            }

            // 若使用 wono 或 eng_sr 查詢，回傳合計數量供前端顯示
            if (!string.IsNullOrWhiteSpace(wono) || !string.IsNullOrWhiteSpace(eng_sr))
            {
                int totalQty = finalResult.Sum(r => r.quantity ?? 0);
                ViewBag.TotalCount = totalQty;
            }

            return View("StoreHouseDropshipping", "_LayoutMember", finalResult);
        }

        /// <summary>
        /// 預計直出資料建立 - View
        /// /// </summary>
        /// <returns></returns>
        public ActionResult CreateDropship()
        {
            // 權限不足或未登入
            var authResult = Login_Authentication();
            // 檢查是否已登入，若未登入則導向登入頁
            if (authResult != null)
            {
                return authResult;
            }
            // 建立一個新的 E_Dropshipping 物件作為 View 的模型
            var model = new E_Dropshipping();
            // 回傳 CreateDropship 頁面，並套用 _LayoutMember 版型
            return View("CreateDropship", "_LayoutMember", model);

        }

        /// <summary>
        /// 預計直出資料建立
        /// </summary>
        /// <param name="wono">工單號碼</param> //新增 工單號碼 By 20251007 Jesse
        /// <param name="date">日期</param>
        /// <param name="DN">DN</param>
        /// <param name="eng_sr">機種名稱</param>
        /// <param name="quantity">數量</param>
        /// <param name="freight">貨運行</param>
        /// <returns></returns>
        [HttpPost]
        public ActionResult CreateDropship(string wono, string date, string DN, string eng_sr, int quantity, string freight)
        {
            E_Dropshipping dropship = new E_Dropshipping();
            dropship.wono = wono;
            dropship.date = date;
            dropship.DN = DN;
            dropship.eng_sr = eng_sr;
            dropship.quantity = quantity;
            dropship.freight = freight;
            dropship.checkOK = false;
            db.E_Dropshipping.Add(dropship);
            db.SaveChanges();
            return RedirectToAction("StoreHouseDropshipping");
        }

        /// <summary>
        /// 預計直出資料編輯 - View
        /// </summary>
        /// <param name="sno"></param>
        /// <returns></returns>
        public ActionResult EditDropship(int sno)
        {
            var result = db.E_Dropshipping.Find(sno);
            return View(result);
        }

        /// <summary>
        /// 預計直出資料編輯
        /// By 有編輯發送Email 20251007 Jesse
        /// </summary>
        /// <param name="sno">Key</param>
        /// <param name="date">日期</param>
        /// <param name="DN">DN</param>
        /// <param name="eng_sr">機種名稱</param>
        /// <param name="quantity">數量</param>
        /// <param name="freight">貨運行</param>
        /// <param name="wono">工單號碼</param> //新增 工單號碼 By 20251007 Jesse
        /// <returns></returns>
        [HttpPost]
        public ActionResult EditDropship(int sno, string date, string DN, string eng_sr, int quantity, string freight, string wono)
        {
            var result = db.E_Dropshipping.Find(sno);
            if (result == null)
            {
                // 若找不到資料則導回列表並顯示錯誤
                TempData["ErrorMessage"] = "找不到指定的預計直出紀錄，無法編輯。";
                return RedirectToAction("StoreHouseDropshipping");
            }

            // 更新欄位
            result.date = date;
            result.DN = DN;
            result.eng_sr = eng_sr;
            result.quantity = quantity;
            result.freight = freight;
            result.wono = wono;

            try
            {
                db.SaveChanges();

                // 因目前不需用，已改用排程發送Email-[WarehouseExpDirectShipExportExcdlSendEmail] By 20251022 Jesse 注釋掉的寄信功能
                #region 寄送編輯通知Email
                // 儲存成功後以背景工作方式寄信（非阻斷流程）
                /* try
                 {
                     // 使用快照避免將 EF tracked entity 傳到背景執行（可能造成生命週期/序列化問題）
                     var snapshot = new E_Dropshipping
                     {
                         sno = result.sno,
                         wono = result.wono,
                         date = result.date,
                         DN = result.DN,
                         eng_sr = result.eng_sr,
                         quantity = result.quantity,
                         freight = result.freight
                     };

                     // 使用安全的背景啟動器來捕捉並記錄例外
                     StartBackgroundTask(() => _emailService.SendDropshipEditedEmailAsync(snapshot), $"SendDropshipEditedEmailAsync sno={snapshot.sno}");
                 }
                 catch (Exception emailEx)
                 {
                     // 只記錄錯誤訊息到 TempData，避免影響主要流程
                     TempData["EmailError"] = "編輯預計直出寄送通知Email失敗: " + emailEx.Message;
                 }
                */
                #endregion
            }
            catch (Exception ex)
            {
                TempData["ErrorMessage"] = "儲存預計直出發生錯誤: " + ex.Message;
            }

            return RedirectToAction("StoreHouseDropshipping");
        }

        /// <summary>
        /// 預計直出資料刪除
        /// By 有刪除發送Email 20251007 Jesse
        /// </summary>
        /// <param name="sno">序號</param>
        /// <returns></returns>
        public ActionResult DeleteDropship(int sno)
        {
            var result = db.E_Dropshipping.Find(sno);
            if (result == null)
            {
                TempData["ErrorMessage"] = "找不到指定的預計直出紀錄，無法刪除。";
                return RedirectToAction("StoreHouseDropshipping");
            }

            // 在刪除前擷取一份快照用於寄信內容
            var snapshot = new E_Dropshipping
            {
                sno = result.sno,
                wono = result.wono,
                date = result.date,
                DN = result.DN,
                eng_sr = result.eng_sr,
                quantity = result.quantity,
                freight = result.freight
            };

            try
            {
                db.E_Dropshipping.Remove(result);
                db.SaveChanges();

                // 因目前不需用，已改用排程發送Email-[WarehouseExpDirectShipExportExcdlSendEmail] By 20251022 Jesse 注釋掉的寄信功能
                #region 寄送刪除通知Email
                // 背景寄送刪除通知（非阻斷流程）
                /*try
                {
                    StartBackgroundTask(() => _emailService.SendDropshipDeletedEmailAsync(snapshot), $"SendDropshipDeletedEmailAsync sno={snapshot.sno}");
                }
                catch (Exception bgEx)
                {
                    TempData["EmailError"] = "刪除預計直出通知Email失敗: " + bgEx.Message;
                }*/
                #endregion
            }
            catch (Exception ex)
            {
                TempData["ErrorMessage"] = "刪除預計直出發生錯誤: " + ex.Message;
            }

            return RedirectToAction("StoreHouseDropshipping");
        }

        /// <summary>
        /// 預計直出資料確認是否存在
        /// </summary>
        /// <param name="sno">序號</param>
        /// <returns></returns>
        public ActionResult CheckDropship(int sno)
        {
            var result = db.E_Dropshipping.Find(sno);
            if (result == null)
            {
                // 若找不到資料，回傳友善訊息並導回列表，避免 NullReferenceException
                TempData["ErrorMessage"] = $"找不到序號 {sno} 的預計直出資料。";
                return RedirectToAction("StoreHouseDropshipping");
            }

            try
            {
                result.checkOK = true;
                db.SaveChanges();
                TempData["SuccessMessage"] = "已確認預計直出資料。";
            }
            catch (Exception ex)
            {
                // 記錄錯誤並顯示友善訊息，避免回傳 500
                TempData["ErrorMessage"] = "確認預計直出發生錯誤: " + ex.Message;
            }

            return RedirectToAction("StoreHouseDropshipping");
        }

        /// <summary>
        /// 匯出預計直出資料為 Excel檔案 By 新增 20251008 Jesse
        /// </summary>
        /// <param name="wono">工單號碼</param>
        /// <param name="eng_sr">機種名稱</param>
        /// <param name="startDate">出貨日期起</param>
        /// <param name="endDate">出貨日期迄</param>
        /// <returns></returns>
        public ActionResult ExportDropship(string wono, string eng_sr, string startDate, string endDate)
        {
            // 權限檢查
            var authResult = Login_Authentication();
            if (authResult != null) return authResult;

            // 依照 StoreHouseDropshipping 的完整查詢邏輯重用，確保匯出與畫面結果一致
            List<EDropshippingQDto> finalResult = new List<EDropshippingQDto>();

            try
            {
                finalResult = GetDropshipResults(wono, eng_sr, startDate, endDate);
            }
            catch (Exception ex)
            {
                ModelState.AddModelError("ExportDropship", ex.Message);
                finalResult = new List<EDropshippingQDto>();
            }


            // 使用 NPOI 建立真正的 .xlsx 檔案（記憶體）
            IWorkbook workbook = new XSSFWorkbook();
            var sheet = workbook.CreateSheet("預計直出");

            // 建立字體與樣式：標題粗體、數字格式
            IFont boldFont = workbook.CreateFont();
            boldFont.IsBold = true;

            ICellStyle headerStyle = workbook.CreateCellStyle();
            headerStyle.SetFont(boldFont);

            IDataFormat dataFormat = workbook.CreateDataFormat();
            ICellStyle integerStyle = workbook.CreateCellStyle();
            integerStyle.DataFormat = dataFormat.GetFormat("0"); // 整數格式

            // By 20251023 Jesse 新增欄位名稱[尾數箱數量]與格式
            // 標題列
            var headerRow = sheet.CreateRow(0);
            var headers = new[] { "工單號碼", "日期", "DN", "機種名稱", "數量", "滿箱數", "貨運行", "尾數箱數量", "確認" };
            // 標題不帶診斷資訊，使用原始標題文字
            for (int i = 0; i < headers.Length; i++)
            {
                var cell = headerRow.CreateCell(i);
                cell.SetCellValue(headers[i]);
                cell.CellStyle = headerStyle;
            }

            int rowIndex = 1;
            foreach (var r in finalResult)
            {
                var row = sheet.CreateRow(rowIndex++);
                row.CreateCell(0).SetCellValue(r.wono ?? "");
                row.CreateCell(1).SetCellValue(r.date ?? "");
                row.CreateCell(2).SetCellValue(r.DN ?? "");
                row.CreateCell(3).SetCellValue(r.eng_sr ?? "");

                var cellQty = row.CreateCell(4);
                if (r.quantity.HasValue)
                {
                    cellQty.SetCellType(CellType.Numeric);
                    cellQty.SetCellValue((double)r.quantity.Value);
                    cellQty.CellStyle = integerStyle;
                }
                else
                {
                    cellQty.SetCellValue(string.Empty);
                }

                var cellFull = row.CreateCell(5);
                if (r.Full_Amount.HasValue)
                {
                    cellFull.SetCellType(CellType.Numeric);
                    cellFull.SetCellValue((double)r.Full_Amount.Value);
                    cellFull.CellStyle = integerStyle;
                }
                else
                {
                    cellFull.SetCellValue(string.Empty);
                }

                row.CreateCell(6).SetCellValue(r.freight ?? "");

                var numberOfTailBoxes = row.CreateCell(7);
                // 處理 NumberOfTailBoxes 為 string 型別
                if (!string.IsNullOrEmpty(r.NumberOfTailBoxes))
                {
                    if (int.TryParse(r.NumberOfTailBoxes, out var tailBoxNum))
                    {
                        numberOfTailBoxes.SetCellType(CellType.Numeric);
                        numberOfTailBoxes.SetCellValue((double)tailBoxNum);
                        numberOfTailBoxes.CellStyle = integerStyle;
                    }
                    else
                    {
                        // 若不是數字，直接寫入原始字串
                        numberOfTailBoxes.SetCellValue(r.NumberOfTailBoxes);
                    }
                }
                else
                {
                    numberOfTailBoxes.SetCellValue(string.Empty);
                }

                row.CreateCell(8).SetCellValue(r.checkOK.GetValueOrDefault() ? "Y" : "");
            }

            // 自動調整欄寬：根據實際標題數目自動調整每一欄
            int columnCount = headers.Length;
            for (int i = 0; i < columnCount; i++)
            {
                try
                {
                    sheet.AutoSizeColumn(i);
                }
                catch
                {
                    // 某些環境下 AutoSizeColumn 可能失敗，忽略避免中斷匯出
                }
            }

            // 凍結標題列（第一列）
            // colSplit = 0 (不凍結左側欄), rowSplit = 1 (在第一列下方凍結)
            sheet.CreateFreezePane(0, 1);

            string fileName = "預計直出.xlsx";
            using (var stream = new MemoryStream())
            {
                workbook.Write(stream);
                var content = stream.ToArray();
                return File(content, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
            }
        }

        /// <summary>
        /// 取得預計直出查詢結果（已去重，先取得 checkOK=1 再 checkOK=0，以保留已確認優先）
        /// By 增加查詢條件及顯示欄[工單號碼]、[滿箱數] 20251007 Jesse
        /// </summary>
        /// <param name="wono">工單號碼（模糊查詢）</param>
        /// <param name="eng_sr">機種名稱（模糊查詢）</param>
        /// <param name="startDate">起始日期（yyyy-MM-dd）</param>
        /// <param name="endDate">結束日期（yyyy-MM-dd）</param>
        /// <returns>回傳已處理好的 EDropshippingQDto 集合</returns>
        private List<EDropshippingQDto> GetDropshipResults(string wono, string eng_sr, string startDate, string endDate)
        {
            // SQL 基本查詢語句，查詢 E_Dropshipping 並左連 E_StoreHouseStock_BOS 取得滿箱數
            // 修改[滿箱數]的取得從 Table[出貨毛淨重明細] By 20251022 Jesse
            string baseSql = @" SELECT a.[sno], a.[wono], a.[date], a.[DN], a.[eng_sr], a.[quantity], a.[freight], a.[checkOK], b.[滿箱台數] AS Full_Amount
                                , CASE 
                                -- 條件 1: 檢查 Full_Amount 是否為零或 NULL，避免除以零
                                WHEN ISNULL(b.[滿箱台數], 0) = 0 THEN ''
                                -- 條件 2: 檢查餘數是否為零，若為零則顯示空字串 ''
                                WHEN a.[quantity] % b.[滿箱台數] = 0 THEN ''
                                -- 條件 3: 餘數不為零，則計算並轉換為 nvarchar(6) 輸出                                
                                ELSE CAST(a.[quantity] % b.[滿箱台數] AS nvarchar(6)) END AS NumberOfTailBoxes
                                FROM E_Dropshipping a
                                LEFT JOIN [dbo].[出貨毛淨重明細] b ON b.[機種] = a.eng_sr
                                WHERE 1=1";

            // 取得今天日期字串
            string today = DateTime.Now.ToString("yyyy -MM-dd");

            // 建立 SQL 參數清單的委派，根據查詢條件決定要加入哪些參數
            Func<bool, bool, bool, List<SqlParameter>> buildParams = (includeDateStart, includeDateEnd, includeToday) =>
            {
                var list = new List<SqlParameter>();
                // 若有起始日期則加入 @startDate 參數
                if (includeDateStart && !string.IsNullOrWhiteSpace(startDate)) list.Add(new SqlParameter("@startDate", startDate));
                // 若有結束日期則加入 @endDate 參數
                if (includeDateEnd && !string.IsNullOrWhiteSpace(endDate)) list.Add(new SqlParameter("@endDate", endDate));
                // 若有工單號碼則加入 @wono 參數
                if (!string.IsNullOrWhiteSpace(wono)) list.Add(new SqlParameter("@wono", wono + "%"));
                // 若有機種名稱則加入 @eng_sr 參數
                if (!string.IsNullOrWhiteSpace(eng_sr)) list.Add(new SqlParameter("@eng_sr", eng_sr + "%"));
                // 若查詢今天則加入 @today 參數
                if (includeToday) list.Add(new SqlParameter("@today", today));
                return list;
            };

            // 建立 SQL 查詢語句與參數的委派，根據 checkOK 狀態與是否使用日期區間決定查詢內容
            Func<int, bool, Tuple<string, SqlParameter[]>> buildQuery = (checkOKValue, useDateRangeParam) =>
            {
                var sb = new System.Text.StringBuilder(baseSql);

                // 若使用日期區間查詢則加入起始與結束日期條件
                if (useDateRangeParam)
                {
                    if (!string.IsNullOrWhiteSpace(startDate)) sb.Append(" AND a.date >= @startDate");
                    if (!string.IsNullOrWhiteSpace(endDate)) sb.Append(" AND a.date <= @endDate");
                }
                else
                {
                    // 若 checkOK=1 則查詢今天
                    if (checkOKValue == 1) sb.Append(" AND a.date = @today");
                }

                // 加入 checkOK 狀態條件
                sb.Append(" AND a.checkOK = " + checkOKValue);

                // 若有工單號碼則加入模糊查詢條件
                if (!string.IsNullOrWhiteSpace(wono)) sb.Append(" AND a.wono LIKE @wono");
                // 若有機種名稱則加入模糊查詢條件
                if (!string.IsNullOrWhiteSpace(eng_sr)) sb.Append(" AND a.eng_sr LIKE @eng_sr");

                // 依日期排序
                sb.Append(" ORDER BY a.date");

                // 根據查詢條件建立參數陣列
                var parameters = useDateRangeParam
                    ? buildParams(!string.IsNullOrWhiteSpace(startDate), !string.IsNullOrWhiteSpace(endDate), false).ToArray()
                    : buildParams(false, false, checkOKValue == 1).ToArray();

                return Tuple.Create(sb.ToString(), parameters);
            };

            // 判斷是否使用日期區間查詢
            bool useDateRange = !string.IsNullOrWhiteSpace(startDate) || !string.IsNullOrWhiteSpace(endDate);

            // 先查詢 checkOK=1（已確認）資料
            var q1 = buildQuery(1, useDateRange);
            var checkedList = db.Database.SqlQuery<EDropshippingQDto>(q1.Item1, q1.Item2).ToList();

            // 再查詢 checkOK=0（未確認）資料
            var q0 = buildQuery(0, useDateRange);
            var uncheckedList = db.Database.SqlQuery<EDropshippingQDto>(q0.Item1, q0.Item2).ToList();

            // 合併已確認與未確認資料，並以工單、日期、機種、DN、數量去重
            var finalResult = checkedList.Concat(uncheckedList)
                .GroupBy(r => new { r.wono, r.date, r.eng_sr, r.DN, r.quantity })
                .Select(g => g.First())
                .ToList();

            // 回傳查詢結果
            return finalResult;
        }

        /// <summary>
        /// 啟動背景工作並於失敗時記錄錯誤至 /Log/email_errors.log。
        /// By 新增記錄發送Email失敗記錄 20251008 Jesse
        /// </summary>
        /// <param name="taskFunc">要執行的非同步工作委派。</param>
        /// <param name="contextMessage">錯誤記錄時的額外訊息（可選）。</param>
        /// <remarks>
        /// 此方法會以 Task.Run 啟動背景工作，若發生例外，將錯誤詳細資訊（包含檔名與行號）寫入 /Log/email_errors.log 檔案，
        /// 並確保記錄失敗不影響主流程。適用於寄送 Email 等非阻斷性通知用途。
        /// </remarks>
        private void StartBackgroundTask(Func<Task> taskFunc, string contextMessage = null)
        {
            try
            {
                var task = System.Threading.Tasks.Task.Run(taskFunc);
                task.ContinueWith(t =>
                {
                    try
                    {
                        // 儲存至 /Log/email_errors.log
                        var basePath = Server != null ? Server.MapPath("~/Log") : Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Log");
                        System.IO.Directory.CreateDirectory(basePath);
                        var logPath = System.IO.Path.Combine(basePath, "email_errors.log");
                        var agg = t.Exception?.Flatten();
                        var exText = agg != null ? agg.ToString() : t.Exception?.ToString();

                        // 嘗試擷取第一個 stack frame 的檔名與行號（若有 PDB）
                        string locationInfo = null;
                        try
                        {
                            var first = agg?.InnerExceptions?.FirstOrDefault() ?? t.Exception;
                            if (first != null)
                            {
                                var st = new StackTrace(first, true);
                                var frame = st.GetFrames()?.FirstOrDefault();
                                if (frame != null)
                                {
                                    var file = frame.GetFileName();
                                    var line = frame.GetFileLineNumber();
                                    if (!string.IsNullOrEmpty(file) && line > 0)
                                        locationInfo = $" at {file}:{line}";
                                }
                            }
                        }
                        catch { }

                        var msg = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {contextMessage ?? "BackgroundTask"} - Exception:{locationInfo}\n{exText}\n";
                        System.IO.File.AppendAllText(logPath, msg);
                    }
                    catch
                    {
                        // 記錄失敗不應影響主流程
                    }
                }, System.Threading.Tasks.TaskContinuationOptions.OnlyOnFaulted);
            }
            catch
            {
                // 啟動背景工作失敗也不應阻斷主流程
            }
        }
        #endregion

        //public GetQCData()
        //{
        //    //取得會員帳號指定fUserId

        //    QC_Model qc = new QC_Model();
        //    var models = qc.GetData();
        //    int SumProCount = models.Sum(s => s.ProCount);
        //    ViewBag.Total = SumProCount;
        //    return View("StoreHouseQCDatas", "_LayoutMember", models);
        //}

        //群組用
        /// <summary>
        /// 依據群組代碼 (Igroup) 查詢庫存明細。
        /// </summary>
        /// <param name="id">群組代碼 (Igroup)</param>
        /// <returns>回傳 StoreHouseStockDetails 的 PartialView，並帶入該群組的庫存資料清單</returns>
        [HttpGet]
        public ActionResult StoreHouseStockDetails(string id)
        {
            var result = db.E_StoreHouseStock.Where(m => m.Igroup == id).ToList();

            return PartialView("StoreHouseStockDetails", result);
        }

        //已銷單庫存Details
        /// <summary>
        /// 取得指定機種名稱與庫位的庫存明細（已銷單庫存明細）。
        /// </summary>
        /// <param name="engsr">機種名稱</param>
        /// <param name="position">庫位</param>
        /// <returns>回傳 StoreHouseStockBOSDetails 的 PartialView，並帶入該機種的庫存資料清單</returns>
        [HttpGet]
        public ActionResult StoreHouseStockBOSDetails(string engsr, string position)
        {
            var result = db.E_StoreHouseStock.Where(m => m.eng_sr == engsr).ToList();

            return PartialView("StoreHouseStockBOSDetails", result);
        }

        #region Login 驗證相關Class
        /// <summary>
        /// 驗證使用者是否已登入，並將使用者資訊寫入 ViewBag，未登入直接導向登入頁。
        /// </summary>
        /// <returns></returns>
        public ActionResult Login_Authentication()
        {
            // 檢查 Session["Member"] 是否存在且型別正確
            if (Session["Member"] is MaintainViewModels member && member != null)
            {
                ViewBag.UserId = member.fUserId;
                ViewBag.RoleId = member.ROLE_ID;
                return null; // 已登入，回傳 null
            }
            // 未登入直接導向登入頁
            return RedirectToAction("Login", "Home");
        }
        #endregion

        /// <summary>
        /// 取得指定公司代碼與聯絡人之公司資訊，並將公司資訊寫入 ViewBag。
        /// </summary>
        /// <param name="code">公司代碼</param>
        /// <param name="contact">聯絡人</param>
        /// <returns>回傳公司資訊頁面 View</returns>
        public ActionResult Consignment(string code, string contact)
        {
            ConsignmentModel consignment = new ConsignmentModel(db, code, contact);

            List<E_Compyany> compList = consignment.CompList();

            foreach (var item in compList)
            {
                ViewBag.comName = item.company_Name;
                ViewBag.contact = item.contact;
                ViewBag.comNo = item.company_No;
                ViewBag.comTel = item.company_Tel;
                ViewBag.extension = item.tel_extension;
                ViewBag.address = item.address;
            }
            return View();
        }
    }
}
