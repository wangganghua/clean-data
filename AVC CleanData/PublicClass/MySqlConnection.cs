using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.IO;
using System.Diagnostics;
using System.Reflection;
using System.Runtime.InteropServices;
using Tamir.SharpSsh.java.lang;

namespace AVC_ClareData
{
    public class MySqlConnection
    { //定义句柄变量
        public static IntPtr hwnd;
        //定义进程ID变量
        public static int pid = 0;
        //获取进程文件id
        [DllImport("User32.dll", CharSet = CharSet.Auto)]
        public static extern int GetWindowThreadProcessId(IntPtr hwnd, out int ID);

        /// <summary>
        /// 连接数据库字符串
        /// </summary>
        private string strSqlConnection = string.Empty;

        public MySqlConnection()
        {
            strSqlConnection = System.Configuration.ConfigurationManager.ConnectionStrings["strSql"].ConnectionString.ToString();//获取配置文件ip、密码ect
        }

        /// <summary>
        /// 返回DataTable数据集
        /// </summary>
        /// <param name="strSql"></param>
        /// <returns></returns>
        public DataTable GetdtTable(string strSql)
        {
            //string file = Application.ExecutablePath;
            // string connectionString =ConfigurationSettings.AppSettings["linqsql.Properties.Settings.AVC_DATAConnectionString"];
            DataTable dt = new DataTable();
            SqlConnection con = new SqlConnection(strSqlConnection);
            try
            {
                con.Open();
                SqlCommand com = new SqlCommand(strSql, con);
                SqlDataAdapter adpt = new SqlDataAdapter(com);
                adpt.Fill(dt);
            }
            catch (SqlException ex) { 
             //   MessageBox.Show(ex.Message); 
                throw ex;
            }
            finally { con.Close(); }
            return dt;
        }

        /// <summary>
        /// 只支持Delete,Update,Insert,Select Count,Select Sum语句,查询后自动关闭连接。
        /// 返回受影响的记录数,返回-1表示出错。
        /// </summary>
        /// <param name="sqlcmd">需要执行的SQL语句</param>
        /// <returns>返回受影响的记录数,返回-1表示出错</returns>
        public int ExecuteNonQuery(string sqlcmd)
        {
            sqlcmd = sqlcmd.Trim();
            if (sqlcmd.Length == 0)
            {
                throw new ArgumentException("SQL命令没有初始化！");
            }
            int result = -1;
            if (!sqlcmd.ToUpper().Trim().StartsWith("DELETE") && !sqlcmd.ToUpper().Trim().StartsWith("UPDATE") && !sqlcmd.ToUpper().Trim().StartsWith("INSERT") &&
                !sqlcmd.ToUpper().Trim().StartsWith("TRUNCATE") && !sqlcmd.ToUpper().Trim().StartsWith("SELECT COUNT") && !sqlcmd.ToUpper().Trim().StartsWith("SELECT SUM") &&
                !sqlcmd.ToUpper().Trim().StartsWith("CREATE TABLE") && !sqlcmd.ToUpper().Trim().StartsWith("CREATE VIEW"))
            {
                throw new ArgumentException("ExecuteNonQuery方法不支持该SQL命令！");
            }
            bool succeeded = false;
            int errorId = 0;
            while (!succeeded)
            {//初始化连接
                SqlConnection scon = new SqlConnection(strSqlConnection);
                try
                {
                    scon.Open();
                    SqlCommand scom = scon.CreateCommand();
                    scom.CommandText = sqlcmd;
                    scom.CommandTimeout =2*600;
                    if (sqlcmd.ToUpper().StartsWith("SELECT COUNT") || sqlcmd.ToUpper().StartsWith("SELECT SUM"))
                    {
                        object obj = scom.ExecuteScalar();
                        if (obj != DBNull.Value)
                            result = (int)obj;
                        else
                            result = 0;
                    }
                    else
                        result = scom.ExecuteNonQuery();
                    succeeded = true;
                }
                catch (SqlException sex)
                {
                    errorId = sex.Number;
                    if (errorId != 1205)
                    {
                        //Debug.WriteLine(sex);
                        throw sex;
                    }
                    Thread.Sleep(3000);
                }
                catch (Exception ex)
                {
                    //throw ex;
                    Debug.WriteLine(ex.Message);
                }
                finally
                {
                    scon.Close();
                }
            }
            return result;
        }


        /// <summary>
        /// 导出excel数据文件
        /// </summary>
        /// <param name="dataTable">datatable数据集</param>
        /// <param name="path">导出路径</param>
        /// <param name="message">返回信息</param>
        public void DataExportToFile(DataTable dataTable, string path, string message = "数据导出完成")
        {
            try
            {
                StreamWriter Writer = new StreamWriter(path, false, Encoding.GetEncoding("gb2312"));
                StringBuilder Builder = new StringBuilder();
                for (int k = 0; k < dataTable.Columns.Count; k++)
                   Builder.Append(dataTable.Columns[k].ColumnName.ToString() + "\t");
                Builder.Append(Environment.NewLine);
                for (int i = 0; i < dataTable.Rows.Count; i++)
                {
                    for (int j = 0; j < dataTable.Columns.Count; j++)
                        Builder.Append(dataTable.Rows[i][j].ToString() + "\t");
                    Builder.Append("\r\n");
                    //Builder.Append(Environment.NewLine);
                }
                Writer.Write(Builder.ToString());
                Writer.Flush();
                Writer.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        /// <summary>
        /// 导出数据到excel
        /// </summary>
        /// <param name="dtTc">DataTable 数据集</param>
        /// <param name="path">保存路径</param>
        /// <param name="tableName">名称</param>
        public void DataOfGetExcel(DataTable dtTc, string path)
        {
            int FormatNum;//保存excel文件格式
            string Version;//excel版本号
            Microsoft.Office.Interop.Excel.Application Application = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook workbook = (Microsoft.Office.Interop.Excel.Workbook)Application.Workbooks.Add(Missing.Value);//激活工作簿
            //Microsoft.Office.Interop.Excel.Worksheet worksheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Worksheets.Add(true);//给工作簿添加一个sheet
            Microsoft.Office.Interop.Excel.Worksheet worksheet = (Microsoft.Office.Interop.Excel.Worksheet)Application.Workbooks[1].Worksheets[1];//给工作簿添加一个sheet
        
            Version = Application.Version;//获取你使用的excel 的版本号
            if (Convert.ToDouble(Version) < 12)//You use Excel 97-2003
            {
                FormatNum = -4143;
            }
            else//you use excel 2007 or later
            {
                FormatNum = 56;
            }
            //生成字段名称
            int k = 0;
            for (int i = 0; i < dtTc.Columns.Count; i++)
            {
                //  Application.Cells[1, k + 1] = dtTc.Columns[i].ColumnName;
                worksheet.Cells.NumberFormatLocal = "@";//
                worksheet.Cells[1, k + 1] = "" + dtTc.Columns[i].ColumnName;
                k++;
            }
            //  //填充数据
            int r = 0, c = 0;
            for (int i = 0; i < dtTc.Rows.Count; i++)
            {
                c = 0;
                for (int j = 0; j < dtTc.Columns.Count; j++)
                {
                    if (dtTc.Rows[i][j].GetType() == typeof(string))
                    {
                        worksheet.Cells.NumberFormatLocal = "@";//设置单元格为文本
                        worksheet.Cells[r + 2, c + 1] = "" + dtTc.Rows[i][j];
                    }
                    else
                    {
                        //Application.Cells[r + 2, c + 1] = dtTc.Rows[i][j];
                        worksheet.Cells.NumberFormatLocal = "@";
                        worksheet.Cells[r + 2, c + 1] = "" + dtTc.Rows[i][j];
                    }
                    c++;
                }
                r++;
            }
            //保存excel
            workbook.SaveAs(path, FormatNum);
            ////关闭excel进程           
            try
            {
                workbook.Close();
                Application.Quit();
                if (Application != null)
                {
                    //获取Excel App的句柄
                    hwnd = new IntPtr(Application.Hwnd);
                    //通过Windows API获取Excel进程ID
                    GetWindowThreadProcessId(hwnd, out pid);
                    if (pid > 0)
                    {
                        Process process = Process.GetProcessById(pid);
                        process.Kill();
                    }
                }
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
            }

        }

        /// <summary>
        /// 写入txt文件数据
        /// </summary>
        /// <param name="datatable"></param>
        /// <param name="path"></param>
        public void DataExportToTextFile(DataTable datatable, string path)
        {
            string debugErrorMessage = "";
          
            for (int i = 0; i < datatable.Columns.Count; i++)//标题
            {
                debugErrorMessage += datatable.Columns[i].ToString() + "\t";
            }
            debugErrorMessage += "\r\n";
            for (int i = 0; i < datatable.Rows.Count; i++)//内容
            {
                for (int a = 0; a < datatable.Columns.Count; a++)
                {
                    debugErrorMessage += datatable.Rows[i][a] + "\t";
                }
                debugErrorMessage += "\r\n";
            }
            File.AppendAllText(path, debugErrorMessage);
        }
    }
}
