using System.Data;
using System.Data.OleDb;
using System.Data.Odbc;
using System.Data.SqlClient;
using MySql.Data.MySqlClient;

namespace libraryLiankun
{
    public class Database
    {
        public enum databaseType
        {
            sqlServer,
            mySQL,
            oracle,
            excel2003,
            excel2007,
            access2003
        }
        public string connectString;
        public databaseType dbt;

        /// <summary>
        /// 设置数据库连接信息（复杂）
        /// </summary>
        /// <param name="dbt">选择连接的数据库种类</param>
        /// <param name="sourceName">主机名称或主机地址</param>
        /// <param name="databaseName">数据库名称</param>
        /// <param name="userName">用户名</param>
        /// <param name="password">数据库密码</param>
        public Database(databaseType dbt, string sourceName, string databaseName, string userName, string password)
        {
            this.dbt = dbt;
            switch (this.dbt)
            {
                case databaseType.sqlServer:
                    connectString = " Data Source=" + sourceName + "; " +
                        "Initial Catalog=" + databaseName + "; " +
                        "User Id=" + userName + "; " +
                        "Password=" + password + "; ";
                    break;
                case databaseType.mySQL:
                    connectString = "Server = " + sourceName + "; " +
                        "Database =" + databaseName + "; " +
                        "Uid = " + userName + "; " +
                        "Pwd = " + password + "; ";
                    break;
                case databaseType.oracle:
                    connectString = "";
                    break;
                default:
                    connectString = null;
                    break;
            }
        }
        /// <summary>
        /// 设置数据库连接信息（简易）
        /// </summary>
        /// <param name="dbt">选择连接的数据库种类</param>
        /// <param name="databaseName">数据库名称或文件路径</param>
        /// <param name="password">数据库密码</param>
        public Database(databaseType dbt, string databaseName, string password = "")
        {
            switch(dbt)
            {
                case databaseType.excel2003:
                    connectString = "Provider=Microsoft.Jet.OLEDB.4.0;" +
                        "Data Source=" + databaseName + ";" +
                        "Extended Properties='Excel 8.0;" +
                        "HDR=Yes;" +
                        "IMEX=2;'";
                    break;
                case databaseType.excel2007:
                    connectString = "Provider=Microsoft.ACE.OLEDB.12.0;" +
                        "Data Source=" + databaseName + ";" +
                        "Extended Properties='Excel 12.0;" +
                        "HDR=Yes;" +
                        "IMEX=2;'";
                    break;
                case databaseType.access2003:
                    if (password == "")
                    {
                        connectString = "Provider = Microsoft.Jet.OleDb.4.0;" +
                            "Data Source = " + databaseName + ";";
                    }
                    else
                    {
                        connectString = "Provider=Microsoft.Jet.OleDb.4.0;" +
                            "Data Source =" + databaseName + ";" +
                            "Jet OLEDB:Database Password = " + password + " ; ";
                    }
                    break;
                default:
                    connectString = null;
                    break;
            }
        }

        /// <summary>
        /// 查询数据库获取到DataTable
        /// </summary>
        /// <param name="sql">sql语句</param>
        /// <returns>返回DataTable</returns>
        public DataTable GetDataTable(string sql)
        {
            if (dbt == databaseType.sqlServer)
            {
                SqlConnection con = new SqlConnection(connectString);
                SqlCommand cmd = new SqlCommand(sql, con);
                DataTable dt = new DataTable();
                con.Open();
                dt.Load(cmd.ExecuteReader());
                con.Close();
                return dt;
            }
            else if (dbt == databaseType.access2003 ||
                dbt == databaseType.excel2003 ||
                dbt == databaseType.excel2007)
            {
                OleDbConnection con = new OleDbConnection(connectString);
                OleDbCommand cmd = new OleDbCommand(sql, con);
                DataTable dt = new DataTable();
                con.Open();
                dt.Load(cmd.ExecuteReader());
                con.Close();
                return dt;
            }
            else if(dbt == databaseType.mySQL)
            {
                MySqlConnection con = new MySqlConnection(connectString);
                MySqlCommand cmd = new MySqlCommand(sql, con);
                DataTable dt = new DataTable();
                con.Open();
                dt.Load(cmd.ExecuteReader());
                con.Close();
                return dt;
            }
            else
                return null;
        }

        /// <summary>
        /// 编辑数据库
        /// </summary>
        /// <param name="sql">sql语句</param>
        /// <returns>受影响的行数</returns>
        public int EditDatabase(string sql)
        {
            if (dbt == databaseType.sqlServer)
            {
                SqlConnection con = new SqlConnection(connectString);
                SqlCommand cmd = new SqlCommand(sql, con);
                con.Open();
                int i = cmd.ExecuteNonQuery();
                con.Close();
                return i;
            }
            else if (dbt == databaseType.access2003 ||
                dbt == databaseType.excel2003 ||
                dbt == databaseType.excel2007)
            {
                OleDbConnection con = new OleDbConnection(connectString);
                OleDbCommand cmd = new OleDbCommand(sql, con);
                con.Open();
                int i = cmd.ExecuteNonQuery();
                con.Close();
                return i;
            }
            else if (dbt == databaseType.mySQL)
            {
                MySqlConnection con = new MySqlConnection(connectString);
                MySqlCommand cmd = new MySqlCommand(sql, con);
                con.Open();
                int i = cmd.ExecuteNonQuery();
                con.Close();
                return i;
            }
            else
                return -1;
        }

    }













}
