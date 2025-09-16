using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using WebStoreHouse.Models;
using WebStoreHouse.ViewModels;
using PagedList;

namespace WebStoreHouse.Controllers
{
    public class HomeController : Controller
    {
        E_StoreHouseEntities db = new E_StoreHouseEntities();
        public ActionResult Login()
        {
            return View();
        }
        /// <summary>
        /// 登入
        /// </summary>
        /// <param name="fUserId">帳號</param>
        /// <param name="fPwd">密碼</param>
        /// <returns></returns>
        //POST:Home/Login
        [HttpPost]
        public ActionResult Login(string fUserId, string fPwd)
        {
            //一帳密取得會員並指定給member
            try
            {
                //var member = db.E_Member.Where(m => m.fUserId == fUserId && m.fPwd == fPwd).FirstOrDefault();
                using (var context = new E_StoreHouseEntities())
                {
                    var member = (from a in context.E_Member
                                  join b in context.E_MemberRole on a.fUserId equals b.USER_ID into roleJoin
                                  from c in roleJoin.DefaultIfEmpty()
                                  where a.fUserId == fUserId && a.fPwd == fPwd
                                  orderby a.fUserId
                                  select new MaintainViewModels
                                  {
                                      fUserId = a.fUserId,
                                      fName = a.fName,
                                      ROLE_ID = c != null ? c.ROLE_ID : null
                                  }).FirstOrDefault();

                    //若member為null表示尚未註冊
                    if (member == null)
                    {
                        ViewBag.Message = "帳號密碼錯誤，請重新登入";
                        return View();
                    }
                    //使用session變數記錄歡迎詞
                    Session["WelCome"] = "員工 : " + member.fName;
                    //使用session變數紀錄登入會員物件
                    Session["Member"] = member;                    
                }
            }
            catch (Exception m)
            {

                throw;
            }

            //執行Home控制器的Index動作
            return RedirectToAction("Index");
        }
        //GET:Home/Logout
        public ActionResult Logout()
        {
            Session.Abandon();
            return RedirectToAction("Index");
        }
        /// <summary>
        /// 首頁-曲線圖
        /// </summary>
        /// <returns></returns>
        public ActionResult Index()
        {
            #region 達交率
            // 取得最新的14筆資料
            List<E_StoreChartjs> dataPoints = db.E_StoreChartjs.OrderByDescending(x => x.sno)
                                                .Take(14).OrderBy(x => x.sno).ToList();
            //TODO:建立X軸用以統計月份            
            string[] days = new string[14];
            DateTime dateF;
            for (int i = 0; i < dataPoints.Count; i++)
            {
                dateF = Convert.ToDateTime(dataPoints[i].date);
                days[i] = dateF.ToString("MM/dd");
            }
            ViewBag.DaysLabel = days;
            int[] data = new int[14];
            for (int j = 0; j < dataPoints.Count; j++)
            {
                data[j] = (int)dataPoints[j].OFR;
            }
            List<ModelChartJs> Dev = new List<ModelChartJs>();
            ModelChartJs model = new ModelChartJs
            {
                OFR = data
            };
            Dev.Add(model);            
            #endregion
            //若session["Member"]=null 表示會員未登入
            if (Session["Member"] == null)
            {                
                //指定Index.cshtml套用_layout.cshtml, View 使用products                
                return View("Index", "_Layout", Dev);
            }
            //會員登入狀態
            //指定Index.cshtml套用_layoutMember.cshtml, View 使用products                        
            return View("Index", "_LayoutMember", Dev);
        }        
        int pagesize = 50;
        public ActionResult Invoicing(int page = 1)
        {
            int currentPage = page < 1 ? 1 : page;
            InvoicingViewModel IVM = new InvoicingViewModel() 
            {
                invoicing = db.Invoicing.ToList()
            };            
            var result = IVM.invoicing.ToPagedList(currentPage, pagesize);
            //var Invo = db.Invoicing.ToList();
            return View(result);
        }
        public ActionResult About()
        {
            ViewBag.Message = "Your application description page.";

            return View();
        }

        public ActionResult Contact()
        {
            ViewBag.Message = "Your contact page.";

            return View();
        }
        public ActionResult ErrDevBar()
        {
            //TODO:建立X軸用以統計月份
            //string[] Months = { "1月", "2月", "3月", "4月", "5月", "6月", "7月", "8月", "9月", "10月", "11月", "12月" };
            //ViewBag.MonthLabel = Months;
            //List<ModelChartJs> Dev = new List<ModelChartJs>
            //{
            //    new ModelChartJs
            //    {
            //        dev="1號機",
            //        errCount=new int[]
            //        {
            //            1,3,5,7,9,12,20,9,10,14,19,20
            //        }
            //    },
            //    new ModelChartJs
            //    {
            //        dev="2號機",
            //        errCount=new int[]
            //        {
            //            1,2,9,8,7,4,1,2,3,6,4,5
            //        }
            //    }
            //};
            return View();
        }
        
    }
}