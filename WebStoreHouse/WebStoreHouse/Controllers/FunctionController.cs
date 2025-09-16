using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using WebStoreHouse.Models;
using WebStoreHouse.ViewModels;
using PagedList;
using System.Diagnostics;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.IO;
using System.Globalization;
using System.Data.Entity;
using static WebStoreHouse.Models.QC_Model;
using static WebStoreHouse.Models.ConsignmentModel;

namespace WebStoreHouse.Controllers
{
    public class FunctionController : Controller
    {
        E_StoreHouseEntities db = new E_StoreHouseEntities();
        int pagesize = 50;
        // GET: Function
        public ActionResult Index()
        {
            return View();
        }
        //view
        public  ActionResult StoreHouseStock(string wono,string order_cust,string engsr, int page = 1)
        {
            int currentPage = page < 1 ? 1 : page;
            int _c = 0;
            //取得會員帳號指定fUserId
            if (!Login_Authentication()){
                return RedirectToAction("Login", "Home");}
            //string UserId = (Session["Member"] as MaintainViewModels).fUserId;
            //string RoleId = (Session["Member"] as MaintainViewModels).ROLE_ID;
            List<StoreHouseStock> stores = new List<StoreHouseStock>();
            List<SqlParameter> parmL = new List<SqlParameter>();
            string strsql = @"select serialno,sno,nowono,wono,cust_wono,eng_sr,order_count,quantity,box_quantity,kf10,kq30,
                                            sap_in,position,acc_in,outed,notout,borrow,due_date,mark,inputdate,package,output_local,Igroup 
                                            from E_StoreHouseStock where  quantity >0 and (del_flag is null or del_flag <> 'D') and (igroup ='' or Igroup is null)  ";
            if (!string.IsNullOrEmpty(wono))
            {
                strsql += " and wono = @wono ";
                parmL.Add(new SqlParameter("wono", wono));
            }
            if (!string.IsNullOrEmpty(order_cust))
            {
                strsql += " and cust_wono = @order_cust ";
                parmL.Add(new SqlParameter("order_cust", order_cust));
            }
            if (!string.IsNullOrEmpty(engsr))
            {
                strsql += " and eng_sr = @engsr ";
                parmL.Add(new SqlParameter("engsr", engsr));
            }
            strsql += @" union
                                select serialno,sno,nowono,wono,cust_wono,eng_sr,order_count,quantity,box_quantity,kf10,kq30,
                                sap_in,position,acc_in,outed,notout,borrow,due_date,mark,inputdate,package,output_local,Igroup 
                                from (Select *,ROW_NUMBER() Over (Partition By Igroup Order By inputdate Desc) As Sort 
                                From E_StoreHouseStock where igroup <>'' and igroup is not null and (del_flag is null or del_flag <> 'D') 
                                and quantity >0) s where sort=1 ";
            if (!string.IsNullOrEmpty(wono))
            {
                strsql += " and wono = @wono2 ";
                parmL.Add(new SqlParameter("wono2", wono));
                _c++;
            }
            if (!string.IsNullOrEmpty(order_cust))
            {
                strsql += " and cust_wono = @order_cust2 ";
                parmL.Add(new SqlParameter("order_cust2", order_cust));
                _c++;
            }
            if (!string.IsNullOrEmpty(engsr))
            {
                strsql += " and eng_sr = @engsr2 ";
                parmL.Add(new SqlParameter("engsr2", engsr));
                _c++;
            }
            strsql += @" order by eng_sr ";

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
            //ViewBag.Member = UserId;
            //ViewBag.RoleId = RoleId;
            currentPage = (_c != 0) ? 1 : currentPage;

            return View("StoreHouseStock", "_LayoutMember", stores.ToPagedList(currentPage, pagesize));
        }
        //create
        public ActionResult StockCreate()
        {
            Session["IGroup"] = string.Empty;
            if (string.IsNullOrEmpty(Session["Member"].ToString()))
            {
                return View();
            }
            return View("StockCreate", "_LayoutMember");
        }
        [HttpPost]
        public JsonResult CheckData(string input)
        {
            bool foundData = db.E_ZRSD19.Any(o => o.wono == input);
            TempData["CheckData"] = foundData ? "" : "not found";
            return Json(new { result = foundData ? "" : "not found" });
        }
        [HttpPost]
        public ActionResult StockCreate(string wono, int? quantity,string position,string mark,string package,
            string buttonName,bool isActive,int box_quantity, string transportation, string nowono="N" )
        {
            #region ZRSD19查無資料無法新增情況
            if (TempData["CheckData"] != null && !string.IsNullOrEmpty(TempData["CheckData"].ToString()))
            {
                ViewBag.err = "ZRSD19查無資料，無法新增!";
                return View("StockCreate", "_LayoutMember");
            }
            #endregion
            if (buttonName == "Group")
            {
                Session["IGroup"] = Guid.NewGuid().ToString();
                TempData["IGroup"] = Session["IGroup"];
                return View();
            }
            if (buttonName == "unGroup")
            {
                Session["IGroup"] = "";
                TempData["IGroup"] = "";
                return View();
            }

            //string Isession = string.IsNullOrEmpty(Session["IGroup"].ToString()) ? "" : Session["IGroup"].ToString();
            if (buttonName != "Group" && isActive==false)
            {
                if (string.IsNullOrEmpty(quantity.ToString()))
                {
                    ViewBag.err = "數量錯誤!";
                    return View();
                }
                #region step 1. 寫入庫存表 In 作業
                E_StoreHouseStock_In stock_In = new E_StoreHouseStock_In();
                stock_In.nowono = nowono;
                stock_In.wono = wono;
                stock_In.quantity = quantity;
                stock_In.box_quantity = box_quantity;
                stock_In.position = position;
                stock_In.mark = mark;
                stock_In.package = package;
                stock_In.transportation = transportation;
                stock_In.inputdate = DateTime.Now.ToString("yyyy-MM-dd");
                stock_In.Igroup = Session["IGroup"].ToString();
                db.E_StoreHouseStock_In.Add(stock_In);
                db.SaveChanges();

                #endregion

                #region step 2. 寫入庫存總表作業
                //找到剛剛寫入_in的sno
                int sno = (from E_StoreHouseStock_In in db.E_StoreHouseStock_In
                           where E_StoreHouseStock_In.wono == wono
                           orderby E_StoreHouseStock_In.sno descending
                           select E_StoreHouseStock_In.sno).FirstOrDefault();
                //透過這筆sno的wono資料查詢相關資訊
                var query = from a in db.E_ZRSD19
                            join b in ((from E_Kf10 in db.E_Kf10 group E_Kf10 by new
                             {E_Kf10.Material} into g
                             select new
                             {
                                 g.Key.Material,
                                 Unrestricted = (int?)g.Sum(p => p.Unrestricted)
                             })) on new { Item = a.item } equals new { Item = b.Material } into b_join
                            from b in b_join.DefaultIfEmpty()
                            join c in ((from E_KQ30 in db.E_KQ30 group E_KQ30 by new
                            {E_KQ30.Material} into g
                             select new
                             {
                                 g.Key.Material,
                                 Unrestricted = (int?)g.Sum(p => p.Unrestricted)
                             })) on new { Item = a.item } equals new { Item = c.Material } into c_join
                            from c in c_join.DefaultIfEmpty()
                            where
                              a.wono == wono
                            select new
                            {
                                a.wono_cust,
                                a.item,
                                a.order_quantity,
                                kf10 =b.Unrestricted == null ? 0 : b.Unrestricted,
                                kq30 =c.Unrestricted == null ? 0 : c.Unrestricted,
                                a.wono_inStoreCount,
                                a.shipped_quantity,
                                a.unshipped_quantity,
                                a.borrow_count,
                                a.due_date,
                            };

                //將該筆wono資料寫入庫存表
                //寫入前確認是否重複 ; stockCount > 0 代表庫存數++
                var stockCount = db.E_StoreHouseStock
                    .Where(stock => stock.wono == wono && stock.position == position && stock.del_flag != "D");
                if (stockCount.Count() > 0)
                {
                    var tstock = db.E_StoreHouseStock.Where(w => w.wono == wono && w.position == position && w.del_flag != "D").FirstOrDefault();
                    tstock.quantity += quantity;
                    tstock.box_quantity += box_quantity;
                }
                else
                {
                    foreach (var q in query)
                    {
                        E_StoreHouseStock houseStock = new E_StoreHouseStock
                        {
                            sno = sno,
                            nowono = nowono,
                            wono = wono,
                            cust_wono = q.wono_cust,
                            eng_sr = q.item,
                            order_count = q.order_quantity,
                            quantity = quantity,
                            box_quantity = box_quantity, //null ? 0 : b.Unrestricted,
                            kf10 = q.kf10,
                            kq30 = q.kq30,
                            position = position,
                            acc_in = q.wono_inStoreCount, //庫存
                            outed = q.shipped_quantity,
                            notout = q.unshipped_quantity,
                            borrow = q.borrow_count,
                            due_date = q.due_date,
                            mark = mark,
                            inputdate = DateTime.Now.ToString("yyyy-MM-dd"),
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
                catch (Exception h)
                {

                }
                #endregion

                //如果按下再一次按鈕則返回新增畫面
                if (buttonName == "Again")
                {
                    ModelState.Clear();
                    return View();
                }
            }
            #region isActive 無單號新增
            if (isActive)
            {
                #region step 1. 寫入庫存表 In 作業
                E_StoreHouseStock_In stock_In = new E_StoreHouseStock_In();
                stock_In.nowono = "Y";
                stock_In.wono = wono;
                stock_In.quantity = quantity;
                stock_In.box_quantity = box_quantity;
                stock_In.position = position;
                stock_In.mark = mark;
                stock_In.package = package;
                stock_In.transportation = transportation;
                stock_In.inputdate = DateTime.Now.ToString("yyyy-MM-dd");
                //stock_In.Igroup = Session["IGroup"].ToString();
                try
                {
                    db.E_StoreHouseStock_In.Add(stock_In);
                    db.SaveChanges();
                }
                catch (Exception chkIN)
                {
                    throw;
                }
                #endregion

                //var chk_in = db.E_StoreHouseStock_In.Where(_in => _in.wono == wono).FirstOrDefault();
                var chk_in = db.E_StoreHouseStock_In.Where(_in => _in.wono == wono)
                    .OrderByDescending(_in => _in.sno).FirstOrDefault();
                E_StoreHouseStock houseStock = new E_StoreHouseStock
                {
                    sno = chk_in.sno,
                    nowono=chk_in.nowono,
                    wono = chk_in.wono,
                    quantity = chk_in.quantity,
                    box_quantity=chk_in.box_quantity,
                    position = chk_in.position,
                    mark = chk_in.mark,
                    inputdate = DateTime.Now.ToString("yyyy-MM-dd"),
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
                catch (Exception h)
                { }
            }

            #endregion
            //如果按下Create按鈕則回到庫存畫面
            return RedirectToAction("StoreHouseStock");
        }
        public ActionResult DeleteStock(int? serialno, string igroup)
        {
            //刪除非真正刪除資料,加入flag以後查詢
            //var stock = db.E_StoreHouseStock.Find(sno);
            List<E_StoreHouseStock> stock;
            if (!string.IsNullOrEmpty(igroup))
            {
                stock = db.E_StoreHouseStock.Where(m => m.Igroup == igroup).ToList();
            }
            else
            {
                stock = db.E_StoreHouseStock.Where(m => m.serialno == serialno).ToList();
            }

            foreach (var stocks in stock)
            {
                stocks.del_flag = "D";
            }

            db.SaveChanges();
            return RedirectToAction("StoreHouseStock");
        }
        public ActionResult EditStock(int serialno)
        {
            var result = db.E_StoreHouseStock.SingleOrDefault(m => m.serialno == serialno);
            return View("EditStock", "_LayoutMember", result);
        }
        [HttpPost]
        public ActionResult EditStock(int serialno, string position,int quantity,string mark,string wono,string oldP,string oldM)
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
                .Where(s =>s.wono== wono && s.position == position && s.del_flag != "D").FirstOrDefault();
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
        public ActionResult AddCar(int sno)
        {
            //取得會員帳號指定給fUserId
            string UserId = (Session["Member"] as MaintainViewModels).fUserId;
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

            //try
            //{
            //    db.SaveChanges();
            //}
            //catch (Exception sc)
            //{

            //    throw;
            //}

            //執行Home控制器的ShoppingCar動作方法
            return RedirectToAction("StoreHouseStock_SC");

        }
        public ActionResult Delete_SC(int sno)
        {
            var del_sc = db.E_StoreHouseStock_SC.Find(sno);
            db.E_StoreHouseStock_SC.Remove(del_sc);
            db.SaveChanges();
            return RedirectToAction("StoreHouseStock_SC");
        }
        //GET:Index/DeleteCar
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
        //GET:Function/StoreHouseStock_SC
        public ActionResult StoreHouseStock_SC(int page = 1)
        {
            int currentPage = page < 1 ? 1 : page;
            //取得會員帳號指定fUserId
            //string UserId = (Session["Member"] as MaintainViewModels).fUserId;
            //string RoleId = (Session["Member"] as MaintainViewModels).ROLE_ID;
            //ViewBag.Member = UserId;
            //ViewBag.RoleId = RoleId;
            if (!Login_Authentication()){
                return RedirectToAction("Login", "Home");}
            //找出未成為訂單明細資料,即購物車內容
            string today = DateTime.Now.ToString("yyyy-MM-dd");
            var orderDetails = db.E_StoreHouseStock_SC.Where
                (m => m.IsApproved == "N" && m.exp_shipdate==today).ToList();
            //指定ShoppingCar.cshtml套用_LayoutMember.cshtml,View使用orderDetail模型
            return View("StoreHouseStock_SC", "_LayoutMember",orderDetails.ToPagedList(currentPage, pagesize));
        }

        /*
         * 此動作會先新增tOrder訂單主檔,接著將tOrderDetail訂單明細的購物車"狀態"之產品的fApproved屬性設為"是"
         * 表示該筆產品正式成為訂單的產品之一, 訂單處理完成後執行Home/OrderList動作方法切換到訂單顯示作業
         */
        //POST:Function/StoreHouseStock_SC
        [HttpPost]
        public ActionResult StoreHouseStock_SC(string fReceiver, string fEmail, string fAddress)
        {
            //找出會員帳號指定給fUserId
            string UserId = (Session["Member"] as MaintainViewModels).fUserId;
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
            var carList = db.E_StoreHouseStock_SC.Where(m => m.IsApproved == "N" ).ToList();
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
            catch (Exception w)
            {

                throw;
            }

            //執行Home控制器的OrderList動作方法
            return RedirectToAction("OrderList");
        }
        public CVM13_Stock Edit_View(int _sno,string _sales_order)
        {
            // 找出要編輯的E_StoreHouseStock_SC資料
            CVM13_Stock CVM13 = new CVM13_Stock()
            {
                //找出未成為訂單明細資料,即購物車內容
                storeHouseStock_SCs=db.E_StoreHouseStock_SC.Where(m=>m.sno==_sno).ToList(),
                //storeHouseStock_SCs = new List<E_StoreHouseStock_SC>() { stockSC },
                storeHouseStocks = db.E_StoreHouseStock
                .Where(m => (m.wono).Substring(0, 7) == _sales_order && m.del_flag != "D" && m.quantity > 0).ToList()
            };
            return CVM13;
        }
        public ActionResult Edit_SC(int sno,string sales_order)
        {
            Edit_View(sno, sales_order);

            return View("Edit_SC", "_LayoutMember", Edit_View(sno, sales_order));
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Edit_SC(int sno, string sales_order,string cust_wono,int? exp_shipquantity,string sap_mark,
            int? stock_quantity,string sap_in,string company_code,string wono,string position,string eng_sr,int ssno,int box_quantity,string transportation)
        {
            //取得會員帳號指定fUserId
            string UserId = (Session["Member"] as MaintainViewModels).fUserId;

            //清除公司代碼
            if (company_code=="")
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
                            stock_Order.inputdate = sdate;
                            stock_Order.position = position;
                            stock_Order.transportation = transportation;
                        }

                        try
                        {
                            db.SaveChanges();
                        }
                        catch (Exception editsc)
                        {
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
                        return View("Edit_SC", "_LayoutMember", Edit_View(ssno, sales_order));
                    }
                }
            }
            return RedirectToAction("StoreHouseStock_SC");

        }
        #region 無單號出貨
        public ActionResult Edit_SCbos(int sno, string sales_order)
        {
            //取得會員帳號指定fUserId
            string UserId = (Session["Member"] as MaintainViewModels).fUserId;

            var bos_sc = db.E_StoreHouseStock_SC.Where(b => b.sno == sno).FirstOrDefault();

            return View("Edit_SCbos", "_LayoutMember", bos_sc);
        }

        [HttpPost]
        public ActionResult Edit_SCbos(int sno, int? exp_shipquantity, string sap_mark,
            int? stock_quantity, string company_code, string position, string eng_sr,string transportation)
        {
            //取得會員帳號指定fUserId
            string UserId = (Session["Member"] as MaintainViewModels).fUserId;

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
                    .Where(s => s.sno == sno ).FirstOrDefault();

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
                catch (Exception editsc)
                {
                    throw;
                }

                //已經在表E_StoreHouseStock_BOS扣完帳
            }
            return RedirectToAction("StoreHouseStock_SC");

        }
        #endregion
        //AJAX 更新庫位資訊
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
        //AJAX 更新工單資訊
        public ActionResult Edit_Update(int sno)
        {
            //取得會員帳號指定fUserId
            //string UserId = (Session["Member"] as MaintainViewModels).fUserId;
            var stocks = db.E_StoreHouseStock.Where(m => m.sno == sno && m.del_flag !="D").FirstOrDefault();
            if (stocks == null)
            {
                return HttpNotFound();
            }
          return  Json(stocks, JsonRequestBehavior.AllowGet);
        }
        public ActionResult StoreHouseStock_Order()
        {
            // 找出會員帳號指定給fUserId
            //string fUserId = (Session["Member"] as MaintainViewModels).fUserId;
            //ViewBag.Member = fUserId;
            if (!Login_Authentication())
            {
                return RedirectToAction("Login", "Home");
            }
            string today = DateTime.Now.ToString("yyyy-MM-dd");
            var result = db.E_StoreHouseStock_Order
                .Where(o => o.chk_date.ToString() == today && o.del_flag !="D").OrderBy(o=>o.sap_mark).ToList();
            return View("StoreHouseStock_Order", "_LayoutMember",result);
        }

        public ActionResult DeleteOrder(int sno,string reason,string wono,int quantity,string guid)
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


            var tstock = db.E_StoreHouseStock.Where(w => w.wono == wono && w.position== record.position).FirstOrDefault();
            tstock.quantity +=  quantity;
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
        [HttpPost]
        public ActionResult OrderConfirm(string sno)
        {
            // 找出會員帳號指定給fUserId
            string fUserId = (Session["Member"] as MaintainViewModels).fUserId;
            //建立唯一的識別,並指定給guid變數用來當作訂單編號
            //tOrder的fOrderGuid欄位會關連到tOrderDetail的fOrderGuid欄位
            //形成一對的關係,即一筆訂單資料對應多筆訂單明細(Master-Detail)
            string guid = Guid.NewGuid().ToString();

            //找出目前在預計出貨中"未有識別碼"的資料
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
            catch (Exception tcarList)
            {
                throw;
            }
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

            //建立當日訂單主檔資料
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
                catch (Exception torder)
                {
                    throw;
                }
            }
            db.SaveChanges();

            #region 寫入已銷單庫存
            foreach (var item_5 in carList_5)
            {
                var engsr_bos = db.E_StoreHouseStock_BOS.FirstOrDefault(x => x.eng_sr == item_5.eng_sr);
                if (engsr_bos!=null)
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
                    catch (Exception torder)
                    {
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

        //bill of sale 銷貨單據,銷單
        public ActionResult StoreHouseStock_BOS()
        {
            if (!Login_Authentication()){
                return RedirectToAction("Login", "Home");}
            if (!string.IsNullOrEmpty(Session["Member"].ToString()))
            {
                //string sql = @"select eng_sr,sum(quantity) as quantity,position,transportation,
                //                            Full_Amount,Full_GW,Full_NW from E_StoreHouseStock_BOS 
                //                            where quantity > 0 
                //                            group by eng_sr,position,transportation,Full_Amount,Full_GW,Full_NW ";
                string sql = @"select Distinct * from E_StoreHouseStock_BOS a
                                            left join E_Weight b on a.eng_sr=b.EngSr where a.quantity > 0 ";
                DataSet dataSet = dbMethod.ExecuteDataSet(sql, CommandType.Text, null);
                return View("StoreHouseStock_BOS", "_LayoutMember", dataSet);
            }
            return View();
        }
        public ActionResult BOS_toShip(string eng_sr,string position)
        {
            var _bos = db.E_StoreHouseStock_BOS.Where(m => m.eng_sr == eng_sr && m.position == position).FirstOrDefault();
            return View("BOS_toShip", "_LayoutMember",_bos);
        }
        [HttpPost]
        public ActionResult BOS_toShip(string eng_sr,int ship_quantity,string position,string transportation)
        {
            // 取得會員帳號指定 fUserId
            string UserId = (Session["Member"] as MaintainViewModels).fUserId;

            using (var transaction = db.Database.BeginTransaction())
            {
                try
                {
                    // 找出符合條件的資料
                    var bos = db.E_StoreHouseStock_BOS
                        .Where(m => m.eng_sr == eng_sr && m.position == position)
                        .ToList();

                    if (bos.Count > 0)
                    {
                        string today = DateTime.Now.ToString("yyyy-MM-dd");

                        // 更新 E_StoreHouseStock_BOS 資料
                        foreach (var item in bos)
                        {
                            item.ship_quantity = ship_quantity;
                            item.quantity -= ship_quantity;
                            item.ship_date = today;
                            item.transportation = transportation;
                        }

                        // 儲存異動
                        db.SaveChanges();

                        // 將資料新增到 E_StoreHouseStock_SC 中
                        var tos = bos.Select(item => new E_StoreHouseStock_SC
                        {
                            exp_shipdate = item.ship_date,
                            eng_sr = item.eng_sr,
                            exp_shipquantity = item.ship_quantity,
                            stock_quantity = item.quantity + item.ship_quantity,
                            position = item.position,
                            transportation=item.transportation,
                            IsApproved = "N",
                            sales_order="N/A",
                            cust_wono="N/A",
                            wono= "無工單出貨"
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
        public ActionResult Edit_BOS(string eng_sr)
        {
            //取得會員帳號指定fUserId
            string UserId = (Session["Member"] as MaintainViewModels).fUserId;

            var _bos = db.E_StoreHouseStock_BOS.Where(m => m.eng_sr == eng_sr && m.quantity >0).FirstOrDefault();
            return View("Edit_BOS", "_LayoutMember", _bos);
        }
        [HttpPost]
        public ActionResult Edit_BOS(string eng_sr, string position,string changeP,int quantity)
        {
            //取得會員帳號指定fUserId
            string UserId = (Session["Member"] as MaintainViewModels).fUserId;
            using (var db =new E_StoreHouseEntities())
            {
                using (var transaction = db.Database.BeginTransaction())
                {
                    try
                    {
                        var bos = db.E_StoreHouseStock_BOS.SingleOrDefault(m => m.eng_sr == eng_sr && m.position == changeP);
                        if (bos != null)
                        {
                            bos.quantity += quantity;
                        }
                        else
                        {
                            var e_bos = db.E_StoreHouseStock_BOS.FirstOrDefault(m => m.eng_sr == eng_sr && m.position == position && m.quantity > 0);
                            e_bos.position = changeP;
                            e_bos.date = e_bos.date;
                            e_bos.eng_sr = e_bos.eng_sr;
                            e_bos.quantity = e_bos.quantity;
                            e_bos.Full_Amount = e_bos.Full_Amount;
                            e_bos.Full_GW = e_bos.Full_GW;
                            e_bos.Full_NW = e_bos.Full_NW;
                            e_bos.ship_quantity = e_bos.ship_quantity;
                            e_bos.ship_date = e_bos.ship_date;
                            e_bos.transportation = e_bos.transportation;
                            db.E_StoreHouseStock_BOS.Add(e_bos);
                            //db.SaveChanges();
                            //transaction.Commit();
                        }

                        //異動後原始庫位資料數量=0,並加上註記
                        var u_bos = db.E_StoreHouseStock_BOS.FirstOrDefault(m => m.eng_sr == eng_sr && m.position == position && m.quantity > 0);
                        u_bos.quantity = 0;

                        db.SaveChanges();
                        transaction.Commit();
                    }
                    catch (Exception ex)
                    {
                        transaction.Rollback();
                        //Log error message here
                        throw;
                    }
                }
            }
            return RedirectToAction("StoreHouseStock_BOS");
        }
        public ActionResult StoreHouseStock_OrderData(string indate, string indate2, string engsr
            ,string wono,string mark,string li,string order_cust,string wono_cust)
        {
            if (!Login_Authentication())
            {
                return RedirectToAction("Login", "Home");
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
            DataSet StockData=oDM.Gds();

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
        public ActionResult StoreHouseDropshipping(string date)
        {
            //取得會員帳號指定fUserId
            if (!Login_Authentication()){
                return RedirectToAction("Login", "Home");}
            // 查詢所有的 E_Dropshipping 資料，或者根據日期篩選資料
            string today = DateTime.Now.ToString("yyyy-MM-dd");

            //找出今天資料&打勾
            var result_1 = db.E_Dropshipping.Where(d => d.date == today && d.checkOK == true && (string.IsNullOrEmpty(date) || d.date == date))
                                          .OrderBy(d => d.date)
                                          .ToList();
            //找出沒有打勾資料
            var result_2 = db.E_Dropshipping.Where(c=>c.checkOK == false && (string.IsNullOrEmpty(date) || c.date == date))
                                          .OrderBy(c => c.date)
                                          .ToList();
            var combinedResult = result_1.Concat(result_2).ToList();
            return View("StoreHouseDropshipping", "_LayoutMember", combinedResult);
        }
        public ActionResult CreateDropship()
        {
            if (!string.IsNullOrEmpty(Session["Member"].ToString()))
            {
                return View("CreateDropship", "_LayoutMember");
            }
            return View();
        }
        [HttpPost]
        public ActionResult CreateDropship(string date, string DN, string eng_sr, int quantity, string freight)
        {
            E_Dropshipping dropship = new E_Dropshipping();
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
        public ActionResult EditDropship(int sno)
        {
            var result = db.E_Dropshipping.Find(sno);
            return View(result);
        }
        [HttpPost]
        public ActionResult EditDropship(int sno,string date,string DN,string eng_sr,int quantity,string freight)
        {
            var result = db.E_Dropshipping.Find(sno);
            result.date = date;
            result.DN = DN;
            result.eng_sr = eng_sr;
            result.quantity = quantity;
            result.freight = freight;
            db.SaveChanges();

            return RedirectToAction("StoreHouseDropshipping");
        }

        public ActionResult CheckDropship(int sno)
        {
            var result = db.E_Dropshipping.Find(sno);
            result.checkOK = true;
            db.SaveChanges();
            return RedirectToAction("StoreHouseDropshipping");
        }

        public ActionResult DeleteDropship(int sno)
        {
            var result = db.E_Dropshipping.Find(sno);
            db.E_Dropshipping.Remove(result);
            db.SaveChanges();

            return RedirectToAction("StoreHouseDropshipping");
        }
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
        [HttpGet]
        public ActionResult StoreHouseStockDetails(string id)
        {
            var result = db.E_StoreHouseStock.Where(m => m.Igroup == id).ToList();

            return PartialView("StoreHouseStockDetails", result);
        }
        //已銷單庫存Details
        [HttpGet]
        public ActionResult StoreHouseStockBOSDetails(string engsr,string position)
        {
            var result = db.E_StoreHouseStock.Where(m => m.eng_sr == engsr ).ToList();

            return PartialView("StoreHouseStockBOSDetails", result);
        }

        #region Login 驗證相關Class
        public bool Login_Authentication()
        {
            if (Session["Member"]  != null)
            {
                string UserId = (Session["Member"] as MaintainViewModels).fUserId;
                string RoleId = (Session["Member"] as MaintainViewModels).ROLE_ID;
                ViewBag.RoleId = RoleId;
                ViewBag.UserId = UserId;
                return true;
            }
            else
            {
                return false;
            }
        }
        #endregion

        #region Excel ActionFunction
        // 在 Controller 中新增一個匯出 Excel 的 Action
        public ActionResult ExportToExcel()
        {
            string today = DateTime.Now.ToString("yyyy-MM-dd");
            // 取得出貨明細資料
            var orders = db.E_StoreHouseStock_Order.Where(o => o.chk_date.ToString() == today).ToList();

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
            headerRow.CreateCell(6).SetCellValue("成倉數量");
            //headerRow.CreateCell(7).SetCellValue("SAP入庫");
            headerRow.CreateCell(7).SetCellValue("公司名稱");
            headerRow.CreateCell(8).SetCellValue("地址");

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
                dataRow.CreateCell(6).SetCellValue(order.stock_quantity.ToString());
                //dataRow.CreateCell(7).SetCellValue(order.sap_in.ToString());
                dataRow.CreateCell(7).SetCellValue(order.company_name);
                dataRow.CreateCell(8).SetCellValue(order.address);
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

        public ActionResult Consignment(string code,string contact)
        {
            ConsignmentModel consignment = new ConsignmentModel(db,code,contact);

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
