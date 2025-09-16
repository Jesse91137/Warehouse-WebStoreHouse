using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Web;

namespace WebStoreHouse.Models
{
    public class OrderData_Model
    {
        public DataSet ds = new DataSet();
        public string indate { get; set; }
        public string indate2 { get; set; }
        public string engsr { get; set; }
        public string wono { get; set; }
        public string docu { get; set; }
        public string mark { get; set; }
        public string li { get; set; }
        public string order_cust { get; set; }
        public string wono_cust { get; set; }

        #region MyRegion
        public DataSet Gds()
        {
            List<SqlParameter> parmL = new List<SqlParameter>();
            string strsql = string.Empty;
            if (li == "O")
            {
                //List<SqlParameter> parmL = new List<SqlParameter>();
                strsql = @"select chk_date,wono,eng_sr,sales_order,cust_wono,exp_shipquantity,
                                            stock_quantity,sap_mark,company_name,address 
                                            from E_StoreHouseStock_Order where wono is not null ";
                //只有開始日有值
                if (!string.IsNullOrEmpty(indate) && string.IsNullOrEmpty(indate2))
                {
                    strsql += @" and chk_date = @indate ";
                    parmL.Add(new SqlParameter("indate", indate));
                }
                //只有結束日有值
                else if (string.IsNullOrEmpty(indate) && !string.IsNullOrEmpty(indate2))
                {
                    strsql += @" and chk_date = @indate2 ";
                    parmL.Add(new SqlParameter("indate2", indate2));
                }
                //兩個都有值
                else if (!string.IsNullOrEmpty(indate) && !string.IsNullOrEmpty(indate2))
                {
                    strsql += @" and chk_date between @indate and @indate2 ";
                    parmL.Add(new SqlParameter("indate", indate));
                    parmL.Add(new SqlParameter("indate2", indate2));
                }

                if (!string.IsNullOrEmpty(engsr))
                {
                    strsql += @" and eng_sr = @engsr ";
                    parmL.Add(new SqlParameter("engsr", engsr));
                }
                if (!string.IsNullOrEmpty(wono))
                {
                    strsql += @" and wono = @wono ";
                    parmL.Add(new SqlParameter("wono", wono));
                }
                if (!string.IsNullOrEmpty(mark))
                {
                    strsql += @" and sap_mark = @mark ";
                    parmL.Add(new SqlParameter("mark", mark));
                }
                //ds = dbMethod.ExecuteDataSetPmsList(strsql, CommandType.Text, parmL);
            }
            else if (li == "Z")
            {
                strsql = @"select top(50)* from E_ZRSD19 where sno is not null ";
                if (!string.IsNullOrEmpty(wono))
                {
                    strsql += @" and wono like @wono+'%' ";
                    parmL.Add(new SqlParameter("wono", wono));
                }
                if (!string.IsNullOrEmpty(engsr))
                {
                    strsql += @" and item = @engsr ";
                    parmL.Add(new SqlParameter("engsr", engsr));
                }
                if (!string.IsNullOrEmpty(order_cust))
                {
                    strsql += @" and order_cust = @order_cust ";
                    parmL.Add(new SqlParameter("order_cust", order_cust));
                }
                if (!string.IsNullOrEmpty(wono_cust))
                {
                    strsql += @" and wono_cust = @wono_cust ";
                    parmL.Add(new SqlParameter("wono_cust", wono_cust));
                }                
            }
            else if (li == "K")
            {
                strsql = @"select distinct b.SLoc,b.Material,b.Docu,b.Unrestricted
                                    from E_StoreHouseStock a
                                    join E_Kf10 b on substring(a.wono,1,7) = b.docu
                                    where substring(a.wono,1,7) in 
                                    (select docu from E_Kf10 where len(docu)>1 group by docu having count(docu)>=2) 
                                    and quantity>0 and del_flag is null 
                                    union
                                    select distinct b.SLoc,b.Material,b.Docu,b.Unrestricted
                                    from E_StoreHouseStock a
                                    join E_KQ30 b on substring(a.wono,1,7) = b.docu
                                    where substring(a.wono,1,7) in 
                                    (select docu from E_KQ30 where len(docu)>1 group by docu having count(docu)>=2) 
                                    and quantity>0 and del_flag is null ";

                if (!string.IsNullOrEmpty(docu))
                {
                    strsql += @" and item = @engsr ";
                    parmL.Add(new SqlParameter("engsr", engsr));
                }
                if (!string.IsNullOrEmpty(order_cust))
                {
                    strsql += @" and order_cust = @order_cust ";
                    parmL.Add(new SqlParameter("order_cust", order_cust));
                }
                if (!string.IsNullOrEmpty(wono_cust))
                {
                    strsql += @" and wono_cust = @wono_cust ";
                    parmL.Add(new SqlParameter("wono_cust", wono_cust));
                }
            }
            ds = (!string.IsNullOrEmpty(strsql))?dbMethod.ExecuteDataSetPmsList(strsql, CommandType.Text, parmL):null;
            return ds;
        }
        #endregion
    }
}