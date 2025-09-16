using PagedList;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using WebStoreHouse.Models;
using WebStoreHouse.ViewModels;

namespace WebStoreHouse.Controllers
{
    public class MaintainController : Controller
    {
        E_StoreHouseEntities db = new E_StoreHouseEntities();
        int pageSize = 50;
        // GET: Maintain
        public ActionResult Index()
        {
            return View();
        }
        public ActionResult StoreHouseStock_Maintain(string li,int page = 1)
        {
            int currentPage = page < 1 ? 1 : page;
            DataSet ds = new DataSet();

            //取得會員帳號指定fUserId
            if (!Login_Authentication()){
                return RedirectToAction("Login", "Home");}

            ViewBag.List = li;
            if (li == "M")
            {
                //sqlstr = @"select a.fUserId,a.fName,c.ROLE_DESC,b.ROLE_ID from E_Member a
                //                    left join E_MemberRole b on a.fUserId=b.USER_ID
                //                    left join E_StoreHouse_Role c on b.ROLE_ID=c.ROLE_ID
                //                    order by fUserId";
                //ds = dbMethod.ExecuteDataSet(sqlstr, CommandType.Text, null);
                using (var context = new E_StoreHouseEntities()) // Replace "YourDbContext" with the actual name of your DbContext class
                {
                    var query = from a in context.E_Member
                                join b in context.E_MemberRole on a.fUserId equals b.USER_ID into roleJoin
                                from c in roleJoin.DefaultIfEmpty()
                                join d in context.E_StoreHouse_Role on c.ROLE_ID equals d.ROLE_ID into roleDescJoin
                                from e in roleDescJoin.DefaultIfEmpty()
                                orderby a.fUserId
                                select new MaintainViewModels
                                {
                                    fId=a.fId,
                                    fUserId=a.fUserId,
                                    fName=a.fName,
                                    ROLE_DESC = e != null ? e.ROLE_DESC : null,
                                    ROLE_ID = c != null ? c.ROLE_ID : null
                                };
                    var result = query.ToPagedList(currentPage, pageSize); // Replace "10" with the desired page size
                    return View("StoreHouseStock_Maintain", "_LayoutMember",result);
                }
            }
            else if (li == "C")
            {
                //sqlstr = @"select* from E_Compyany";
                //ds = dbMethod.ExecuteDataSet(sqlstr, CommandType.Text, null);
                using (var context = new E_StoreHouseEntities()) // Replace "YourDbContext" with the actual name of your DbContext class
                {
                    var query = from f in context.E_Compyany.OrderBy(x=>x.sno) // Typo: should be "E_Company"?
                                select f;
                    var result = query.ToPagedList(currentPage, pageSize); // Replace "10" with the desired page size
                    return View("StoreHouseStock_Maintain", "_LayoutMember", result);
                }
            }
            else
            {
                return View("StoreHouseStock_Maintain", "_LayoutMember");
            }
        }
        public ActionResult CreateMember()
        {
            return View("CreateMember", "_LayoutMember");
        }
        [HttpPost]
        public ActionResult CreateMember(string fUserId, string fName,string ROLE_ID)
        {
            var repeat = db.E_Member.Where(u => u.fUserId == fUserId).FirstOrDefault();
            if (repeat != null)
            {
                // 使用TempData儲存資料
                TempData["ErrorMessage"] = "該使用者代號已存在，請輸入不同的代號";
                TempData["fUserId"] = fUserId;
                TempData["fName"] = fName;
                //TempData["ROLE_ID"] = ROLE_ID;
                return View("CreateMember", "_LayoutMember");
            }
            else
            {
                //人員名單
                E_Member member = new E_Member();
                member.fUserId = fUserId;
                member.fPwd = fUserId;
                member.fName = fName;
                db.E_Member.Add(member);
                db.SaveChanges();

                //人員權限名單
                E_MemberRole member_Role = new E_MemberRole();
                member_Role.USER_ID = fUserId;
                member_Role.ROLE_ID = ROLE_ID;
                member_Role.EXPIRED_DATE = DateTime.Now;
                db.E_MemberRole.Add(member_Role);
                db.SaveChanges();
            }

            return RedirectToAction("StoreHouseStock_Maintain", new { li = "M" });
        }
        public ActionResult Edit_Member(int fId)
        {
            using (var context = new E_StoreHouseEntities()) // Replace "YourDbContext" with the actual name of your DbContext class
            {
                var query = from a in context.E_Member
                            join b in context.E_MemberRole on a.fUserId equals b.USER_ID into roleJoin
                            from c in roleJoin.DefaultIfEmpty()
                            join d in context.E_StoreHouse_Role on c.ROLE_ID equals d.ROLE_ID into roleDescJoin
                            from e in roleDescJoin.DefaultIfEmpty()
                            where a.fId == fId
                            orderby a.fUserId
                            select new MaintainViewModels
                            {
                                fId = a.fId,
                                fUserId = a.fUserId,
                                fName = a.fName,
                                ROLE_DESC = e != null ? e.ROLE_DESC : null,
                                ROLE_ID = c != null ? c.ROLE_ID : null
                            };
                var result = query.ToList(); // Replace "10" with the desired page size
                return View("Edit_Member","_LayoutMember",result);
            }
        }
        [HttpPost]
        public ActionResult Edit_Member(string fUserId,string fName,string ROLE_ID)
        {
            //人員名單
            var member = db.E_Member.Where(m => m.fUserId == fUserId).FirstOrDefault();
            member.fUserId = fUserId;
            member.fName = fName;
            db.SaveChanges();

            //人員權限名單
            var member_Role = db.E_MemberRole.Where(r => r.USER_ID == fUserId).FirstOrDefault();
            member_Role.USER_ID = fUserId;
            member_Role.ROLE_ID = ROLE_ID;
            member_Role.EXPIRED_DATE = DateTime.Now;
            db.SaveChanges();

            return RedirectToAction("StoreHouseStock_Maintain", new { li = "M" });
        }
        public ActionResult Del_Member(int fId)
        {
            using (var context = new E_StoreHouseEntities())
            {
                var member = context.E_Member.FirstOrDefault(m => m.fId == fId);
                if (member != null)
                {
                    context.E_Member.Remove(member);

                    var memberRoles = context.E_MemberRole.Where(mr => mr.USER_ID == member.fUserId);
                    if (memberRoles != null)
                    {
                        context.E_MemberRole.RemoveRange(memberRoles);
                    }

                    context.SaveChanges();
                }
            }
            return RedirectToAction("StoreHouseStock_Maintain", new { li = "M" });
        }

        public ActionResult CreateCompInfo()
        {
            return View("CreateCompInfo", "_LayoutMember");
        }
        [HttpPost]
        public ActionResult CreateCompInfo(string code, string company_No, string company_Name
            , string address, string company_Tel, string tel_extension, string contact)
        {
            var repeat = db.E_Compyany.Where(c => c.code == code).FirstOrDefault();
            if (repeat != null)
            {
                // 使用TempData儲存資料
                TempData["ErrorMessage"] = "該代碼已存在，請輸入不同的代碼";
                TempData["code"] = code;
                TempData["company_No"] = company_No;
                TempData["company_Name"] = company_Name;
                TempData["address"] = address;
                TempData["company_Tel"] = company_Tel;
                TempData["tel_extension"] = tel_extension;
                TempData["contact"] = contact;
                return View("CreateCompInfo", "_LayoutMember");
            }
            else
            {
                //人員名單
                E_Compyany comp = new E_Compyany();
                comp.code = code;
                comp.company_No = company_No;
                comp.company_Name = company_Name;
                comp.address = address;
                comp.company_Tel = company_Tel;
                comp.tel_extension = tel_extension;
                comp.contact = contact;
                db.E_Compyany.Add(comp);
                db.SaveChanges();
            }            
            return RedirectToAction("StoreHouseStock_Maintain", new { li = "M" });
        }
        public ActionResult Edit_CompInfo(int sno)
        {
            using (var context = new E_StoreHouseEntities())
                {
                    var query = from f in context.E_Compyany.OrderBy(x=>x.sno).Where(s=>s.sno == sno)
                                select f;
                    var result = query.ToList();
                    return View("Edit_CompInfo", "_LayoutMember",result);
                }
        }
        [HttpPost]
        public ActionResult Edit_CompInfo(int sno,string code,string company_No,string company_Name
            ,string address,string company_Tel,string tel_extension,string contact)
        {
            //公司資料
            var comp = db.E_Compyany.Where(s => s.sno == sno).FirstOrDefault();
            comp.code = code;
            comp.company_No = company_No;
            comp.company_Name = company_Name;
            comp.address = address;
            comp.company_Tel = company_Tel;
            comp.tel_extension = tel_extension;
            comp.contact = contact;

            db.SaveChanges();

            return RedirectToAction("StoreHouseStock_Maintain", new { li = "C" });
        }
        public ActionResult DelCompInfo(int sno)
        {
            var comp = db.E_Compyany.FirstOrDefault(s => s.sno==sno);
            db.E_Compyany.Remove(comp);
            db.SaveChanges();

            return RedirectToAction("StoreHouseStock_Maintain", new { li = "C" });
        }
        #region Login 驗證相關Class
        public bool Login_Authentication()
        {
            if (Session["Member"] != null)
            {
                MaintainViewModels member = Session["Member"] as MaintainViewModels;
                ViewBag.RoleId = member.ROLE_ID;
                ViewBag.UserId = member.fUserId;

                return true;
            }
            else
            {
                return false;
            }
        }
        #endregion
    }
}