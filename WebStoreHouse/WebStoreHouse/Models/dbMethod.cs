using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Web;

namespace WebStoreHouse.Models
{
    public class dbMethod
    {
        #region paramater method

        //開發測試
        //private static readonly String connStr = "server=localhost;database=E_StoreHouse;uid=sa;pwd=A12345678;Connect Timeout = 480";
        //private static readonly String connStr = ConfigurationManager.ConnectionStrings["StoreHouseStock"].ToString();

        //正式環境
        //private static readonly String connStr = "server=192.168.4.120;database=E_StoreHouse;uid=stock_web;pwd=A12345678;Connect Timeout = 480";
        //ConfigurationManager.ConnectionStrings["conString"].ConnectionString;    


        /// <summary>
        /// 
        /// </summary>
        private static readonly string connStr;

        static dbMethod()
        {
            ConnectionStringSettings connStrsettings = ConfigurationManager.ConnectionStrings["StoreHouseStock"];
            if (connStrsettings != null)
            {
                connStr = connStrsettings.ConnectionString;
            }
            else
            {
                // 記錄錯誤或使用預設連接字串
                Console.WriteLine("Error: StoreHouseStock connection string not found.");
                //throw new Exception("StoreHouseStock connection string not found.");
            }
        }


        //1. 執行insert/update/delete，回傳影響的資料列數
        public static int ExecueNonQuery(string sql, CommandType cmdType, params SqlParameter[] pms)
        {
            using (SqlConnection con = new SqlConnection(connStr))
            {
                using (SqlCommand cmd = new SqlCommand(sql, con))
                {
                    //設置目前執行的是「存儲過程? 還是帶參數的sql 語句?」
                    cmd.CommandType = cmdType;
                    if (pms != null)
                    {
                        cmd.Parameters.AddRange(pms);
                    }

                    con.Open();
                    return cmd.ExecuteNonQuery();
                }
            }
        }
        public static SqlDataReader ExecuteReader(string sql, CommandType cmdType, params SqlParameter[] pms)
        {
            SqlConnection con = new SqlConnection(connStr);
            using (SqlCommand cmd = new SqlCommand(sql, con))
            {
                cmd.CommandType = cmdType;
                if (pms != null)
                {
                    cmd.Parameters.AddRange(pms);
                }
                try
                {
                    con.Open();
                    return cmd.ExecuteReader(CommandBehavior.CloseConnection);
                }
                catch
                {
                    con.Close();
                    con.Dispose();
                    throw;
                }
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sql"></param>
        /// <param name="cmdType"></param>
        /// <param name="pms"></param>
        /// <returns></returns>
        public static DataTable ExecuteDataTable(string sql, CommandType cmdType, params SqlParameter[] pms)
        {
            DataTable dt = new DataTable();
            //use SqlDataAdapter ,it will establish Sql connection.So ,it no need to create Connection by yourself.
            using (SqlDataAdapter adapter = new SqlDataAdapter(sql, connStr))
            {
                adapter.SelectCommand.CommandType = cmdType;
                if (pms != null)
                {
                    adapter.SelectCommand.Parameters.AddRange(pms);

                }
                adapter.Fill(dt);
                return dt;
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sql"></param>
        /// <param name="cmdType"></param>
        /// <param name="pms"></param>
        /// <returns></returns>
        public static DataSet ExecuteDataSet(string sql, CommandType cmdType, params SqlParameter[] pms)
        {
            DataSet ds = new DataSet();
            //use SqlDataAdapter ,it will establish Sql connection.So ,it no need to create Connection by yourself.
            using (SqlDataAdapter adapter = new SqlDataAdapter(sql, connStr))
            {
                adapter.SelectCommand.CommandType = cmdType;
                if (pms != null)
                {
                    adapter.SelectCommand.Parameters.AddRange(pms);

                }
                adapter.Fill(ds);
                return ds;
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sql"></param>
        /// <param name="cmdType"></param>
        /// <param name="pms"></param>
        /// <returns></returns>
        public static DataSet ExecuteDataSetPmsList(string sql, CommandType cmdType, List<SqlParameter> pms)
        {
            DataSet ds = new DataSet();
            //use SqlDataAdapter ,it will establish Sql connection.So ,it no need to create Connection by yourself.
            using (SqlDataAdapter adapter = new SqlDataAdapter(sql, connStr))
            {
                adapter.SelectCommand.CommandType = cmdType;
                if (pms != null)
                {
                    adapter.SelectCommand.Parameters.AddRange(pms.ToArray<SqlParameter>());//paralist.ToArray<SqlParameter>()

                }
                adapter.Fill(ds);
                return ds;
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sql"></param>
        /// <param name="cmdType"></param>
        /// <param name="pms"></param>
        /// <returns></returns>
        public static SqlDataReader ExecuteReaderPmsList(string sql, CommandType cmdType, List<SqlParameter> pms)
        {
            SqlConnection con = new SqlConnection(connStr);
            using (SqlCommand cmd = new SqlCommand(sql, con))
            {
                cmd.CommandType = cmdType;
                if (pms != null)
                {
                    cmd.Parameters.AddRange(pms.ToArray<SqlParameter>());
                }
                try
                {
                    con.Open();
                    return cmd.ExecuteReader(CommandBehavior.CloseConnection);
                }
                catch
                {
                    con.Close();
                    con.Dispose();
                    throw;
                }
            }
        }
        #endregion
    }
}