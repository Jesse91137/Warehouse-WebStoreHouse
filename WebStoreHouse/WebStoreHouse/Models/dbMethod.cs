using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Threading.Tasks;
using System.Web;

namespace WebStoreHouse.Models
{
    /// <summary>
    /// 資料庫方法類別，提供與資料庫互動的方法。
    /// </summary>
    public class dbMethod
    {
        #region 連接字串與參數設定

        // 開發測試環境的連接字串 (已註解，可根據需要取消註解)
        // private static readonly string _connStr = "server=localhost;database=E_StoreHouse;uid=sa;pwd=A12345678;Connect Timeout = 480";             

        // 正式環境的連接字串 (已註解，可根據需要取消註解)
        // private static readonly string _connStr = "server=192.168.4.120;database=E_StoreHouse;uid=stock_web;pwd=A12345678;Connect Timeout = 480";
        // private static readonly string _connStr = ConfigurationManager.ConnectionStrings["conString"].ConnectionString;

        /// <summary>
        /// 使用 Lazy<string> 延遲初始化資料庫連接字串。
        /// </summary>
        private static readonly Lazy<string> _connStr = new Lazy<string>(GetConnectionString);

        /// <summary>
        /// 從組態檔中讀取連接字串。
        /// </summary>
        /// <returns>資料庫連接字串。</returns>
        private static string GetConnectionString()
        {
            // 從組態檔中讀取名為 "StoreHouseStock" 的連接字串設定
            ConnectionStringSettings connStrSettings = ConfigurationManager.ConnectionStrings["StoreHouseStock"];

            // 如果找到連接字串設定，則返回其值
            if (connStrSettings != null)
            {
                return connStrSettings.ConnectionString;
            }
            else
            {
                // 如果找不到連接字串設定，則記錄錯誤訊息
                Console.WriteLine("Error: StoreHouseStock connection string not found.");

                //找不到連接字串時，返回null
                return null;
            }
        }

        /// <summary>
        /// 取得資料庫連接字串。
        /// </summary>
        private static string ConnStr => _connStr.Value;

        /// <summary>
        /// 預設的命令超時時間 (秒)。
        /// </summary>
        private const int CommandTimeout = 480; // 8 分鐘

        #endregion

        #region 資料庫操作方法

        /// <summary>
        /// 設定 SqlCommand 的參數和命令類型。
        /// </summary>
        /// <param name="cmd">SqlCommand 物件。</param>
        /// <param name="cmdType">命令類型。</param>
        /// <param name="sql">SQL 查詢字串。</param>
        /// <param name="pms">參數列表。</param>
        private static void ConfigureCommand(SqlCommand cmd, CommandType cmdType, string sql, SqlParameter[] pms)
        {
            cmd.CommandType = cmdType;
            cmd.CommandText = sql;
            cmd.CommandTimeout = CommandTimeout; // 設定命令超時時間

            if (pms != null)
            {
                // 驗證參數是否為 null 或包含 null 值
                if (pms.Any(p => p == null || p.Value == null))
                {
                    throw new ArgumentException("SQL 參數不能為 null 或包含 null 值。");
                }
                cmd.Parameters.AddRange(pms);
            }
        }

        /// <summary>
        /// 執行 SqlCommand 並處理資源釋放。
        /// </summary>
        /// <typeparam name="T">返回類型。</typeparam>
        /// <param name="func">要執行的委派。</param>
        /// <param name="connectionString">資料庫連接字串。</param>
        /// <returns>執行結果。</returns>
        private static T ExecuteCommand<T>(Func<SqlConnection, SqlCommand, T> func, string connectionString)
        {
            // 確保連接字串不為 null
            if (string.IsNullOrEmpty(connectionString))
            {
                throw new InvalidOperationException("資料庫連接字串未設定。");
            }

            using (SqlConnection con = new SqlConnection(connectionString))
            {
                using (SqlCommand cmd = new SqlCommand())
                {
                    cmd.Connection = con;
                    con.Open();
                    return func(con, cmd);
                }
            }
        }

        /// <summary>
        /// 執行 INSERT、UPDATE 或 DELETE 語句，並返回受影響的資料列數。
        /// </summary>
        /// <param name="sql">要執行的 SQL 語句。</param>
        /// <param name="cmdType">命令類型 (Text 或 StoredProcedure)。</param>
        /// <param name="pms">SQL 參數陣列。</param>
        /// <returns>受影響的資料列數。</returns>
        public static int ExecueNonQuery(string sql, CommandType cmdType, params SqlParameter[] pms)
        {
            return ExecuteCommand((con, cmd) =>
            {
                ConfigureCommand(cmd, cmdType, sql, pms);
                return cmd.ExecuteNonQuery();
            }, ConnStr);
        }

        /// <summary>
        /// 執行 INSERT、UPDATE 或 DELETE 語句的非同步版本，並返回受影響的資料列數。
        /// </summary>
        /// <param name="sql">要執行的 SQL 語句。</param>
        /// <param name="cmdType">命令類型 (Text 或 StoredProcedure)。</param>
        /// <param name="pms">SQL 參數陣列。</param>
        /// <returns>受影響的資料列數。</returns>
        public static async Task<int> ExecuteNonQueryAsync(string sql, CommandType cmdType, params SqlParameter[] pms)
        {
            return await ExecuteCommand(async (con, cmd) =>
            {
                ConfigureCommand(cmd, cmdType, sql, pms);
                return await cmd.ExecuteNonQueryAsync();
            }, ConnStr);
        }

        /// <summary>
        /// 在交易中執行 INSERT、UPDATE 或 DELETE 語句，並返回受影響的資料列數。
        /// </summary>
        /// <param name="sql">要執行的 SQL 語句。</param>
        /// <param name="cmdType">命令類型 (Text 或 StoredProcedure)。</param>
        /// <param name="pms">SQL 參數陣列。</param>
        /// <returns>受影響的資料列數。</returns>
        public static int ExecuteNonQueryWithTransaction(string sql, CommandType cmdType, params SqlParameter[] pms)
        {
            return ExecuteCommand((con, cmd) =>
            {
                ConfigureCommand(cmd, cmdType, sql, pms);
                SqlTransaction transaction = con.BeginTransaction();
                cmd.Transaction = transaction;
                try
                {
                    int result = cmd.ExecuteNonQuery();
                    transaction.Commit();
                    return result;
                }
                catch
                {
                    transaction.Rollback();
                    throw;
                }
            }, ConnStr);
        }

        /// <summary>
        /// 執行查詢並返回 SqlDataReader，用於讀取資料。
        /// </summary>
        /// <param name="sql">要執行的 SQL 語句。</param>
        /// <param name="cmdType">命令類型 (Text 或 StoredProcedure)。</param>
        /// <param name="pms">SQL 參數陣列。</param>
        /// <returns>SqlDataReader 物件。</returns>
        public static SqlDataReader ExecuteReader(string sql, CommandType cmdType, params SqlParameter[] pms)
        {
            SqlConnection con = new SqlConnection(ConnStr);
            using (SqlCommand cmd = new SqlCommand())
            {
                ConfigureCommand(cmd, cmdType, sql, pms);
                cmd.Connection = con;
                try
                {
                    con.Open();
                    // 使用 CommandBehavior.CloseConnection 確保在關閉讀取器時關閉連接
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
        /// 執行查詢並返回 SqlDataReader 的非同步版本。
        /// </summary>
        /// <param name="sql">要執行的 SQL 語句。</param>
        /// <param name="cmdType">命令類型 (Text 或 StoredProcedure)。</param>
        /// <param name="pms">SQL 參數陣列。</param>
        /// <returns>SqlDataReader 物件。</returns>
        public static async Task<SqlDataReader> ExecuteReaderAsync(string sql, CommandType cmdType, params SqlParameter[] pms)
        {
            SqlConnection con = new SqlConnection(ConnStr);
            using (SqlCommand cmd = new SqlCommand())
            {
                ConfigureCommand(cmd, cmdType, sql, pms);
                cmd.Connection = con;
                try
                {
                    await con.OpenAsync();
                    // 使用 CommandBehavior.CloseConnection 確保在關閉讀取器時關閉連接
                    return await cmd.ExecuteReaderAsync(CommandBehavior.CloseConnection);
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
        /// 執行查詢並返回 SqlDataReader (使用 List&lt;SqlParameter&gt;)。
        /// </summary>
        /// <param name="sql">要執行的 SQL 語句。</param>
        /// <param name="cmdType">命令類型 (Text 或 StoredProcedure)。</param>
        /// <param name="pms">SQL 參數列表。</param>
        /// <returns>SqlDataReader 物件。</returns>
        public static SqlDataReader ExecuteReaderPmsList(string sql, CommandType cmdType, List<SqlParameter> pms)
        {
            SqlConnection con = new SqlConnection(ConnStr);
            using (SqlCommand cmd = new SqlCommand())
            {
                ConfigureCommand(cmd, cmdType, sql, pms.ToArray());
                cmd.Connection = con;
                try
                {
                    con.Open();
                    // 使用 CommandBehavior.CloseConnection 確保在關閉讀取器時關閉連接                   
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
        /// 執行查詢並返回 SqlDataReader 的非同步版本 (使用 List&lt;SqlParameter&gt;)。
        /// </summary>
        /// <param name="sql">要執行的 SQL 語句。</param>
        /// <param name="cmdType">命令類型 (Text 或 StoredProcedure)。</param>
        /// <param name="pms">SQL 參數列表。</param>
        /// <returns>SqlDataReader 物件。</returns>
        public static async Task<SqlDataReader> ExecuteReaderPmsListAsync(string sql, CommandType cmdType, List<SqlParameter> pms)
        {
            SqlConnection con = new SqlConnection(ConnStr);
            using (SqlCommand cmd = new SqlCommand())
            {
                ConfigureCommand(cmd, cmdType, sql, pms.ToArray());
                cmd.Connection = con;
                try
                {
                    await con.OpenAsync();
                    // 使用 CommandBehavior.CloseConnection 確保在關閉讀取器時關閉連接
                    return await cmd.ExecuteReaderAsync(CommandBehavior.CloseConnection);
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
        /// 執行查詢並返回 DataTable。
        /// </summary>
        /// <param name="sql">要執行的 SQL 語句。</param>
        /// <param name="cmdType">命令類型 (Text 或 StoredProcedure)。</param>
        /// <param name="pms">SQL 參數陣列。</param>
        /// <returns>DataTable 物件。</returns>
        public static DataTable ExecuteDataTable(string sql, CommandType cmdType, params SqlParameter[] pms)
        {
            return ExecuteCommand((con, cmd) =>
            {
                ConfigureCommand(cmd, cmdType, sql, pms);
                DataTable dt = new DataTable();
                using (SqlDataAdapter adapter = new SqlDataAdapter(cmd))
                {
                    adapter.Fill(dt);
                }
                return dt;
            }, ConnStr);
        }

        /// <summary>
        /// 執行查詢並返回 DataTable 的非同步版本。
        /// </summary>
        /// <param name="sql">要執行的 SQL 語句。</param>
        /// <param name="cmdType">命令類型 (Text 或 StoredProcedure)。</param>
        /// <param name="pms">SQL 參數陣列。</param>
        /// <returns>DataTable 物件。</returns>
        public static async Task<DataTable> ExecuteDataTableAsync(string sql, CommandType cmdType, params SqlParameter[] pms)
        {
            return await ExecuteCommand(async (con, cmd) =>
            {
                ConfigureCommand(cmd, cmdType, sql, pms);
                DataTable dt = new DataTable();
                using (SqlDataAdapter adapter = new SqlDataAdapter(cmd))
                {
                    await Task.Run(() => adapter.Fill(dt)); // 使用 Task.Run 執行 Fill
                }
                return dt;
            }, ConnStr);
        }


        /// <summary>
        /// 執行查詢並返回 DataSet。
        /// </summary>
        /// <param name="sql">要執行的 SQL 語句。</param>
        /// <param name="cmdType">命令類型 (Text 或 StoredProcedure)。</param>
        /// <param name="pms">SQL 參數陣列。</param>
        /// <returns>DataSet 物件。</returns>
        public static DataSet ExecuteDataSet(string sql, CommandType cmdType, params SqlParameter[] pms)
        {
            return ExecuteCommand((con, cmd) =>
            {
                ConfigureCommand(cmd, cmdType, sql, pms);
                DataSet ds = new DataSet();
                using (SqlDataAdapter adapter = new SqlDataAdapter(cmd))
                {
                    adapter.Fill(ds);
                }
                return ds;
            }, ConnStr);
        }

        /// <summary>
        /// 非同步執行查詢並返回 DataSet。
        /// </summary>
        /// <param name="sql">要執行的 SQL 語句。</param>
        /// <param name="cmdType">命令類型 (Text 或 StoredProcedure)。</param>
        /// <param name="pms">SQL 參數陣列。</param>
        /// <returns>DataSet 物件。</returns>
        public static async Task<DataSet> ExecuteDataSetAsync(string sql, CommandType cmdType, params SqlParameter[] pms)
        {
            return await ExecuteCommand(async (con, cmd) =>
            {
                ConfigureCommand(cmd, cmdType, sql, pms);
                DataSet ds = new DataSet();
                using (SqlDataAdapter adapter = new SqlDataAdapter(cmd))
                {
                    await Task.Run(() => adapter.Fill(ds)); // 使用 Task.Run 執行 Fill
                }
                return ds;
            }, ConnStr);
        }

        /// <summary>
        /// 執行查詢並返回 DataSet (使用 List&lt;SqlParameter&gt;)。
        /// </summary>
        /// <param name="sql">要執行的 SQL 語句。</param>
        /// <param name="cmdType">命令類型 (Text 或 StoredProcedure)。</param>
        /// <param name="pms">SQL 參數列表。</param>
        /// <returns>DataSet 物件。</returns>
        public static DataSet ExecuteDataSetPmsList(string sql, CommandType cmdType, List<SqlParameter> pms)
        {
            return ExecuteCommand((con, cmd) =>
            {
                ConfigureCommand(cmd, cmdType, sql, pms.ToArray());
                DataSet ds = new DataSet();
                using (SqlDataAdapter adapter = new SqlDataAdapter(cmd))
                {
                    adapter.Fill(ds);
                }
                return ds;
            }, ConnStr);
        }

        /// <summary>
        /// 非同步執行查詢並返回 DataSet (使用 List&lt;SqlParameter&gt;)。
        /// </summary>
        /// <param name="sql">要執行的 SQL 語句。</param>
        /// <param name="cmdType">命令類型 (Text 或 StoredProcedure)。</param>
        /// <param name="pms">SQL 參數列表。</param>
        /// <returns>DataSet 物件。</returns>
        public static async Task<DataSet> ExecuteDataSetPmsListAsync(string sql, CommandType cmdType, List<SqlParameter> pms)
        {
            return await ExecuteCommand(async (con, cmd) =>
            {
                ConfigureCommand(cmd, cmdType, sql, pms.ToArray());
                DataSet ds = new DataSet();
                using (SqlDataAdapter adapter = new SqlDataAdapter(cmd))
                {
                    await Task.Run(() => adapter.Fill(ds)); // 使用 Task.Run 執行 Fill
                }
                return ds;
            }, ConnStr);
        }

        #endregion
    }
}