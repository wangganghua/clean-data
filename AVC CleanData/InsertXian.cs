using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Threading;
using System.Net.NetworkInformation;
using System.Text.RegularExpressions;

namespace AVC_ClareData
{
    public partial class InsertXian : Form
    {
        string scon1 = "Data Source=124.89.13.18,1433;Initial Catalog=dpcdata;User Id=sa;Password=All_View_Consulting_2014@;";
        MySqlConnection mysql = new MySqlConnection();
        public InsertXian()
        {
            InitializeComponent();
            this.progressBar1.Visible = false;
        }
        //型号表、属性表、品牌表 永久表  带机型编码（除永久表）
        private void getTable_sx()
        {
            string begintime = DateTime.Now.ToString();
            //第二批
            string[] category = new[] { "净水器", "净饮机", "饮水机", "净化器", "电熨斗", "吸尘器", "电风扇", "电暖器", "电磁炉", "电蒸炖锅", "除湿机" };
            //第三批
            category = new[] { "干衣机", "彩电", "冰箱", "冰柜", "洗衣机", "智能机顶盒", "投影仪" };
            //category = new[] { "干衣机", "彩电", "冰箱",  "洗衣机", "智能机顶盒", "投影仪" };
            category = new[] { "智能马桶" };
            using (SqlConnection scon = new SqlConnection(scon1))
            {
                scon.Open();
                SqlCommand scom = scon.CreateCommand();
                scom.CommandTimeout = 600;
                SqlDataAdapter sda = new SqlDataAdapter(scom);
                writeR(label1, "开始推送品牌");
                string strsql = "SELECT A.品牌,品牌中文,品牌英文,品牌别名,品牌类型,国别,写入者,写入日期,确认者,确认日期,备注 FROM 品牌表 A LEFT JOIN (SELECT 品牌 FROM OPENROWSET('sqloledb','124.89.13.18,1433';'sa';'All_View_Consulting_2014@',dpcdata.dbo.品牌表))B ON A.品牌=B.品牌 WHERE B.品牌 IS NULL";
                DataTable pingpai = mysql.GetdtTable(strsql);
                if (pingpai.Rows.Count > 0)
                {
                    using (SqlBulkCopy bcp = new SqlBulkCopy(scon))
                    {
                        bcp.DestinationTableName = "品牌表";
                        bcp.BulkCopyTimeout = 600;
                        bcp.BatchSize = 1000;
                        for (int i = 0; i < pingpai.Columns.Count; i++)
                            bcp.ColumnMappings.Add(pingpai.Columns[i].ColumnName.ToString(), pingpai.Columns[i].ColumnName.ToString());
                        bcp.WriteToServer(pingpai);
                        bcp.Close();
                    }
                }
                for (int a = 0; a < category.Length; a++)
                {
                    scom.CommandText = "DELETE 型号表 WHERE 品类='" + category[a] + "'";
                    scom.ExecuteNonQuery();// # 删除型号表该品类数据
                    writeR(label1, "开始推送型号---[" + category[a] + "]");
                    //2、先添加型号表、属性表
                    string sql = " select  a.机型编码,a.品类,a.品牌,a.机型,a.上市月度,a.上市周度,a.子品牌,a.国标机型,a.写入日期,a.写入者,a.审核标记,a.品牌类型细分,a.上市日度 from 型号表 a left join  (SELECT 品牌,品类,机型 FROM OPENROWSET('sqloledb','124.89.13.18,1433';'sa';'All_View_Consulting_2014@',dpcdata.dbo.型号表)) b on a.品类=b.品类 and a.品牌=b.品牌 and a.机型=b.机型 where a.品类='" + category[a] + "' and b.机型 is null";
                    DataTable xihao = mysql.GetdtTable(sql);
                    if (xihao.Rows.Count > 0)
                    {
                        using (SqlBulkCopy bcp = new SqlBulkCopy(scon1, SqlBulkCopyOptions.KeepIdentity | SqlBulkCopyOptions.FireTriggers))
                        {
                            bcp.DestinationTableName = "型号表";
                            bcp.BulkCopyTimeout = 600;
                            bcp.BatchSize = 1000;
                            for (int i = 0; i < xihao.Columns.Count; i++)
                                bcp.ColumnMappings.Add(xihao.Columns[i].ColumnName.ToString(), xihao.Columns[i].ColumnName.ToString());
                            bcp.WriteToServer(xihao);
                            bcp.Close();
                        }
                    }
                    //属性表
                    //计算字段串
                    writeR(label1, "开始推送属性表---[" + category[a] + "]");
                    string jisuanzd = string.Empty;
                    //找出计算字段并剔除
                    sql = "SELECT name FROM sys.computed_columns WHERE object_id = object_id('att" + category[a] + "属性表')";
                    DataTable dt = new DataTable();
                    dt = mysql.GetdtTable(sql);
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        jisuanzd += dt.Rows[i][0].ToString() + ",";
                    }
                    jisuanzd = jisuanzd.TrimEnd(',');
                    //放到一个数组里
                    string[] jsAr = jisuanzd.Split(',');
                    //取属性表字段
                    sql = "SELECT COLUMN_NAME FROM  INFORMATION_SCHEMA.COLUMNS where TABLE_NAME='att" + category[a] + "属性表'  AND COLUMN_NAME NOT IN('ZZID')";
                    DataTable sxdtziduan = mysql.GetdtTable(sql);
                    //字段名串
                    string suxingziduanming = string.Empty;
                    //标识
                    bool t = false;
                    //找出所有不是计算字段的列名
                    for (int i = 0; i < sxdtziduan.Rows.Count; i++)
                    {
                        for (int j = 0; j < jsAr.Length; j++)
                        {
                            t = false;
                            if (jsAr[j] == sxdtziduan.Rows[i][0].ToString())
                            {
                                break;
                            }
                            t = true;
                        }
                        if (t)
                        {
                            if (sxdtziduan.Rows[i][0].ToString() != "zzid")
                            {
                                suxingziduanming += "a.[" + sxdtziduan.Rows[i][0].ToString() + "],";
                            }
                        }
                    }
                    suxingziduanming = suxingziduanming.Replace("a.销额,", "");
                    //删除结尾字符‘，’
                    suxingziduanming = suxingziduanming.TrimEnd(',');//.Replace("a.[机型编码],", "");
                    string[] splitstring = suxingziduanming.Replace("a.", "").Split(',');
                    sql = " select " + suxingziduanming + " from att" + category[a] + "属性表 a left join OPENROWSET('sqloledb','124.89.13.18,1433';'sa';'All_View_Consulting_2014@',dpcdata.dbo.att" + category[a] + "属性表)  b on a.品类=b.品类 and a.品牌=b.品牌 and a.机型=b.机型 where a.品类='" + category[a] + "' and b.机型 is null ";
                    DataTable shuxing = mysql.GetdtTable(sql);
                    if (shuxing.Rows.Count > 0)
                    {
                        //有错，插入时只能插入一条，不可批量，原因不清，2017-08-21 原因：是因为有索引约束
                        using (SqlBulkCopy bcp = new SqlBulkCopy(scon1, SqlBulkCopyOptions.FireTriggers | SqlBulkCopyOptions.KeepIdentity))
                        {
                            bcp.DestinationTableName = "att" + category[a] + "属性表";
                            bcp.BulkCopyTimeout = 1000;
                            bcp.BatchSize = 1000;
                            for (int i = 0; i < shuxing.Columns.Count; i++)
                                bcp.ColumnMappings.Add(splitstring[i], splitstring[i]);
                            bcp.WriteToServer(shuxing);
                            bcp.Close();
                        }
                    }
                    writeR(label1, "" + DateTime.Now.ToString() + "：推送属性表结束---[" + category[a] + "]");
                }
                //3、后添加永久表数据
                string strcategory = string.Empty;
                for (int i = 0; i < category.Length; i++)
                {
                    strcategory += "'" + category[i] + "',";
                }
                strcategory = strcategory.Remove(strcategory.LastIndexOf(","));

                string sc = "SELECT 'SELECT NAME FROM SYSOBJECTS WHERE TYPE=''U'' AND NAME like '''+组别简称+'_%'+品类+'永久表2017年%''  AND NAME NOT LIKE ''%_back'' AND NAME NOT LIKE ''%备份''  and name NOT LIKE ''%2018%''   UNION ALL  ' FROM 品类表 WHERE  品类 IN (" + strcategory + ") ORDER BY 品类";//" + MyConfiguration.UserID + "
                DataTable dtSelect = mysql.GetdtTable(sc);
                string strsxt = string.Empty;
                if (dtSelect.Rows.Count > 0)
                {
                    for (int i = 0; i < dtSelect.Rows.Count; i++)
                    {
                        strsxt += "" + dtSelect.Rows[i][0] + "  ";
                    }
                    strsxt = strsxt.Remove(strsxt.ToUpper().LastIndexOf("UNION ALL"));
                    DataTable dttime = mysql.GetdtTable(strsxt);
                    for (int a = 0; a < dttime.Rows.Count; a++)//永久表
                    {
                        writeR(label2, "共 " + (a + 1) + " / " + dttime.Rows.Count + " 个:");
                        writeR(label1, "" + DateTime.Now.ToString() + "：开始推送永久---[" + dttime.Rows[a]["name"] + "]");
                        //对远程数据库进行操作
                        string sqlcmd = "SELECT COLUMN_NAME FROM  INFORMATION_SCHEMA.COLUMNS where TABLE_NAME='" + dttime.Rows[a]["name"] + "' AND column_name!='zzid'";

                        DataTable dtziduan = mysql.GetdtTable(sqlcmd);
                        //字段名串
                        string ziduanming = string.Empty;
                        for (int i = 0; i < dtziduan.Rows.Count; i++)
                        {
                            ziduanming += dtziduan.Rows[i][0] + ",";
                        }
                        ziduanming = ziduanming.Replace("销额,", "");//.Replace("机型编码,", "");
                        //删除结尾字符‘，’
                        ziduanming = ziduanming.TrimEnd(',');

                        sqlcmd = "SELECT " + ziduanming + " FROM  " + dttime.Rows[a]["name"] + "";
                        DataTable dttable = mysql.GetdtTable(sqlcmd);
                        using (SqlBulkCopy bcp = new SqlBulkCopy(scon))
                        {
                            writeR(label1, "" + DateTime.Now.ToString() + "：开始清空远程永久表数据......]");
                            //清空远程永久表
                            scom.CommandText = "TRUNCATE TABLE " + dttime.Rows[a]["name"] + "";
                            scom.ExecuteNonQuery();
                            bcp.DestinationTableName = dttime.Rows[a]["name"].ToString();
                            bcp.BulkCopyTimeout = 600;
                            bcp.BatchSize = 1000;
                            writeR(label1, "" + DateTime.Now.ToString() + "：开始推送数据：" + bcp.DestinationTableName + "]");
                            for (int i = 0; i < dttable.Columns.Count; i++)
                                bcp.ColumnMappings.Add(dttable.Columns[i].ColumnName.ToString(), dttable.Columns[i].ColumnName.ToString());
                            bcp.WriteToServer(dttable);
                            bcp.Close();
                            writeR(label1, "" + DateTime.Now.ToString() + "：推送结束：" + bcp.DestinationTableName + "]");
                            Debug.WriteLine(DateTime.Now.ToString());
                            Debug.WriteLine(dttime.Rows[a]["name"]);
                        }
                        writeR(label1, "" + DateTime.Now.ToString() + "：开始更新机型编码---[" + dttime.Rows[a]["name"] + "]");
                        scom.CommandText = "update A set 机型编码=b.机型编码  from " + dttime.Rows[a]["name"] + " a join 型号表 b on a.机型=b.机型 and a.品牌=b.品牌 and a.品类=b.品类";
                        scom.ExecuteNonQuery();
                        writeR(label1, "" + DateTime.Now.ToString() + "：机型编码更新结束---[" + dttime.Rows[a]["name"] + "]");
                    }
                }
            }
            string endtime = DateTime.Now.ToString();
            MessageBox.Show("完成: " + begintime + " --- " + endtime);
            return;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.button1.Enabled = false;
            textBox1.Text = DateTime.Now.ToString();
            Thread tr = new Thread(tuiyjb);
            tr.Start();
            //gettest();
            ////Thread tr = new Thread(getTable_sanxing);
            ////tr.Start();
            ////pingIp();
            //test();
        }

        public void writeR(Label c, string aa)
        {
            try
            {
                lock (this)
                {
                    c.Invoke(new ThreadStart(delegate()
                    {
                        c.Text = aa;
                    }));
                }
            }
            catch { }
        }

        //推送 品牌表、型号表、属性表 （left join）
        private void getTable()
        {
            string[] category = new[] { "净水器", "净饮机", "饮水机", "净化器", "电熨斗", "吸尘器", "电风扇", "电暖器", "电磁炉", "电蒸炖锅", "除湿机" };
            //第三批
            category = new[] { "干衣机", "彩电", "冰箱", "冰柜", "洗衣机", "智能机顶盒", "投影仪" };
            using (SqlConnection scon = new SqlConnection(scon1))
            {
                scon.Open();
                SqlCommand scom = scon.CreateCommand();
                scom.CommandTimeout = 600;
                SqlDataAdapter sda = new SqlDataAdapter(scom);
                writeR(label1, "开始推送品牌");
                string strsql = "SELECT A.品牌,品牌中文,品牌英文,品牌别名,品牌类型,国别,写入者,写入日期,确认者,确认日期,备注 FROM 品牌表 A LEFT JOIN (SELECT 品牌 FROM OPENROWSET('sqloledb','124.89.13.18,1433';'sa';'All_View_Consulting_2014@',dpcdata.dbo.品牌表))B ON A.品牌=B.品牌 WHERE B.品牌 IS NULL";
                DataTable pingpai = mysql.GetdtTable(strsql);
                if (pingpai.Rows.Count > 0)
                {
                    using (SqlBulkCopy bcp = new SqlBulkCopy(scon))
                    {
                        bcp.DestinationTableName = "品牌表";
                        bcp.BulkCopyTimeout = 600;
                        bcp.BatchSize = 1000;
                        for (int i = 0; i < pingpai.Columns.Count; i++)
                            bcp.ColumnMappings.Add(pingpai.Columns[i].ColumnName.ToString(), pingpai.Columns[i].ColumnName.ToString());
                        bcp.WriteToServer(pingpai);
                        bcp.Close();
                    }
                }
                for (int a = 0; a < category.Length; a++)
                {
                    writeR(label1, "开始推送型号---[" + category[a] + "]");
                    //2、先添加型号表、属性表
                    string sql = " select  a.品类,a.品牌,a.机型,a.上市月度,a.上市周度,a.子品牌,a.国标机型,a.写入日期,a.写入者,a.审核标记,a.品牌类型细分,a.上市日度 from 型号表 a left join  (SELECT 品牌,品类,机型 FROM OPENROWSET('sqloledb','124.89.13.18,1433';'sa';'All_View_Consulting_2014@',dpcdata.dbo.型号表)) b on a.品类=b.品类 and a.品牌=b.品牌 and a.机型=b.机型 where a.品类='" + category[a] + "' and b.机型 is null";
                    DataTable xihao = mysql.GetdtTable(sql);
                    if (xihao.Rows.Count > 0)
                    {
                        using (SqlBulkCopy bcp = new SqlBulkCopy(scon1, SqlBulkCopyOptions.KeepIdentity | SqlBulkCopyOptions.FireTriggers))
                        {
                            bcp.DestinationTableName = "型号表";
                            bcp.BulkCopyTimeout = 600;
                            bcp.BatchSize = 1000;
                            for (int i = 0; i < xihao.Columns.Count; i++)
                                bcp.ColumnMappings.Add(xihao.Columns[i].ColumnName.ToString(), xihao.Columns[i].ColumnName.ToString());
                            bcp.WriteToServer(xihao);
                            bcp.Close();
                        }
                    }
                    //属性表
                    //计算字段串
                    writeR(label1, "开始推送属性表---[" + category[a] + "]");
                    string jisuanzd = string.Empty;
                    //找出计算字段并剔除
                    sql = "SELECT name FROM sys.computed_columns WHERE object_id = object_id('att" + category[a] + "属性表')";
                    DataTable dt = new DataTable();
                    dt = mysql.GetdtTable(sql);
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        jisuanzd += dt.Rows[i][0].ToString() + ",";
                    }
                    jisuanzd = jisuanzd.TrimEnd(',');
                    //放到一个数组里
                    string[] jsAr = jisuanzd.Split(',');
                    //取属性表字段
                    sql = "SELECT COLUMN_NAME FROM  INFORMATION_SCHEMA.COLUMNS where TABLE_NAME='att" + category[a] + "属性表'  AND COLUMN_NAME NOT IN('ZZID')";
                    DataTable sxdtziduan = mysql.GetdtTable(sql);
                    //字段名串
                    string suxingziduanming = string.Empty;
                    //标识
                    bool t = false;
                    //找出所有不是计算字段的列名
                    for (int i = 0; i < sxdtziduan.Rows.Count; i++)
                    {
                        for (int j = 0; j < jsAr.Length; j++)
                        {
                            t = false;
                            if (jsAr[j] == sxdtziduan.Rows[i][0].ToString())
                            {
                                break;
                            }
                            t = true;
                        }
                        if (t)
                        {
                            if (sxdtziduan.Rows[i][0].ToString() != "zzid")
                            {
                                suxingziduanming += "a.[" + sxdtziduan.Rows[i][0].ToString() + "],";
                            }
                        }
                    }
                    suxingziduanming = suxingziduanming.Replace("a.销额,", "");
                    //删除结尾字符‘，’
                    suxingziduanming = suxingziduanming.TrimEnd(',').Replace("a.[机型编码],", "");
                    string[] splitstring = suxingziduanming.Replace("a.", "").Split(',');
                    sql = " select " + suxingziduanming + " from att" + category[a] + "属性表 a left join OPENROWSET('sqloledb','124.89.13.18,1433';'sa';'All_View_Consulting_2014@',dpcdata.dbo.att" + category[a] + "属性表)  b on a.品类=b.品类 and a.品牌=b.品牌 and a.机型=b.机型 where a.品类='" + category[a] + "' and b.机型 is null ";
                    DataTable shuxing = mysql.GetdtTable(sql);
                    if (shuxing.Rows.Count > 0)
                    {
                        //有错，插入时只能插入一条，不可批量，原因不清，2017-08-21 原因：是因为有索引约束
                        using (SqlBulkCopy bcp = new SqlBulkCopy(scon1, SqlBulkCopyOptions.FireTriggers | SqlBulkCopyOptions.KeepIdentity))
                        {
                            bcp.DestinationTableName = "att" + category[a] + "属性表";
                            bcp.BulkCopyTimeout = 1000;
                            bcp.BatchSize = 1000;
                            for (int i = 0; i < shuxing.Columns.Count; i++)
                                bcp.ColumnMappings.Add(splitstring[i], splitstring[i]);
                            bcp.WriteToServer(shuxing);
                            bcp.Close();
                        }
                    }
                    writeR(label1, "" + DateTime.Now.ToString() + "：推送属性表结束---[" + category[a] + "]");
                }
            }
            MessageBox.Show("完成");
            return;
        }

        //更新永久表机型编码
        private void updatejxbm()
        {
            string begintime = DateTime.Now.ToString();
            string[] category = new[] { "净水器", "净饮机", "饮水机", "净化器", "电熨斗", "吸尘器", "电风扇", "电暖器", "电磁炉", "电蒸炖锅", "除湿机" };
            category = new[] { "电压力锅" };
            category = new[] { "干衣机", "彩电", "冰箱", "冰柜", "洗衣机", "智能机顶盒", "投影仪" };
            using (SqlConnection scon = new SqlConnection(scon1))
            {
                scon.Open();
                SqlCommand scom = scon.CreateCommand();
                scom.CommandTimeout = 600;
                SqlDataAdapter sda = new SqlDataAdapter(scom);
                string strcategory = string.Empty;
                for (int i = 0; i < category.Length; i++)
                {
                    strcategory += "'" + category[i] + "',";
                }
                strcategory = strcategory.Remove(strcategory.LastIndexOf(","));

                string sc = "SELECT 'SELECT NAME FROM SYSOBJECTS WHERE TYPE=''U'' AND NAME like '''+组别简称+'_%'+品类+'永久表____年%''  AND NAME NOT LIKE ''%_back'' AND NAME NOT LIKE ''%备份''  AND  name  LIKE ''%2017%''  AND  name NOT LIKE ''%2018%'' UNION ALL  ' FROM 品类表 WHERE  品类 IN (" + strcategory + ") ORDER BY 品类";//" + MyConfiguration.UserID + "
                DataTable dtSelect = mysql.GetdtTable(sc);
                string strsxt = string.Empty;
                if (dtSelect.Rows.Count > 0)
                {
                    for (int i = 0; i < dtSelect.Rows.Count; i++)
                    {
                        strsxt += "" + dtSelect.Rows[i][0] + "  ";
                    }
                    strsxt = strsxt.Remove(strsxt.ToUpper().LastIndexOf("UNION ALL"));
                    DataTable dttime = mysql.GetdtTable(strsxt);
                    for (int a = 0; a < dttime.Rows.Count; a++)//永久表
                    {
                        writeR(label1, "" + DateTime.Now.ToString() + "：开始更新机型编码---[" + dttime.Rows[a]["name"] + "]");
                        scom.CommandText = "update A set 机型编码=b.机型编码  from " + dttime.Rows[a]["name"] + " a join 型号表 b on a.机型=b.机型 and a.品牌=b.品牌 and a.品类=b.品类";
                        scom.ExecuteNonQuery();
                        writeR(label1, "" + DateTime.Now.ToString() + "：机型编码更新结束---[" + dttime.Rows[a]["name"] + "]");
                    }
                    string endtime = DateTime.Now.ToString();
                    MessageBox.Show("完成: " + begintime + " --- " + endtime);
                    return;
                }
            }
        }

        //推送三星数据
        private void getTable_sanxing()
        {
            string begintime = DateTime.Now.ToString();
            //第二批
            string[] category = new[] { "净水器", "净饮机", "饮水机", "净化器", "电熨斗", "吸尘器", "电风扇", "电暖器", "电磁炉", "电蒸炖锅", "除湿机" };
            //第三批
            category = new[] { "干衣机", "彩电", "冰箱", "冰柜", "洗衣机", "智能机顶盒", "投影仪" };
            category = new[] { "冰箱","洗衣机" };
            category = new[] {"洗衣机" };
            category = new[] { "冰箱" };
            using (SqlConnection scon = new SqlConnection(scon1))
            {
                scon.Open();
                SqlCommand scom = scon.CreateCommand();
                scom.CommandTimeout = 600;
                SqlDataAdapter sda = new SqlDataAdapter(scom);
            
                //3、后添加永久表数据
                string strcategory = string.Empty;
                for (int i = 0; i < category.Length; i++)
                {
                    strcategory += "'" + category[i] + "',";
                }
                strcategory = strcategory.Remove(strcategory.LastIndexOf(","));

                string sc = "SELECT 'SELECT NAME FROM SYSOBJECTS WHERE TYPE=''U'' AND NAME like '''+组别简称+'_%'+品类+'永久表2017年%''  AND NAME NOT LIKE ''%_back'' AND NAME NOT LIKE ''%备份''  and name NOT LIKE ''%2018%''   UNION ALL  ' FROM 品类表 WHERE  品类 IN (" + strcategory + ") ORDER BY 品类";//" + MyConfiguration.UserID + "
                sc = "SELECT 'SELECT NAME FROM SYSOBJECTS WHERE TYPE=''U'' AND (NAME like ''三星'+品类+'%永久表____年%''  or  name like ''三星'+品类+'周度同比库____年'' OR NAME LIKE ''三星'+品类+'属性表'')  UNION ALL  ' FROM 品类表 WHERE  品类 IN (" + strcategory + ") ORDER BY 品类";
                DataTable dtSelect = mysql.GetdtTable(sc);
                string strsxt = string.Empty;
                if (dtSelect.Rows.Count > 0)
                {
                    for (int i = 0; i < dtSelect.Rows.Count; i++)
                    {
                        strsxt += "" + dtSelect.Rows[i][0] + "  ";
                    }
                    strsxt = strsxt.Remove(strsxt.ToUpper().LastIndexOf("UNION ALL"));
                    DataTable dttime = mysql.GetdtTable(strsxt);
                    for (int a = 0; a < dttime.Rows.Count; a++)//永久表
                    {
                        writeR(label2, "共 " + (a + 1) + " / " + dttime.Rows.Count + " 个:");
                        writeR(label1, "" + DateTime.Now.ToString() + "：[三星]开始推送永久---[" + dttime.Rows[a]["name"] + "]");
                        if (a < 10)
                            continue;
                        //1、查找计算列
                        string sql = " SELECT NAME FROM sys.computed_columns WHERE object_id = object_id('" + dttime.Rows[a]["name"] + "') ";
                        DataTable dtcolumn = mysql.GetdtTable(sql);
                        string columnstr = string.Empty;
                        if (dtcolumn.Rows.Count > 0)
                        {
                            for (int c = 0; c < dtcolumn.Rows.Count; c++)
                            {
                                columnstr += "'" + dtcolumn.Rows[c]["name"] + "',";
                            }
                            columnstr = columnstr.Remove(columnstr.LastIndexOf(","));
                        }
                        if (columnstr == "")
                            columnstr = "''";

                        //对远程数据库进行操作
                        string sqlcmd = "SELECT COLUMN_NAME FROM  INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='" + dttime.Rows[a]["name"] + "' AND COLUMN_NAME NOT IN(" + columnstr + ")  AND COLUMN_NAME NOT IN ('ZZID')";

                        DataTable dtziduan = mysql.GetdtTable(sqlcmd);
                        //字段名串
                        string ziduanming = string.Empty;
                        for (int i = 0; i < dtziduan.Rows.Count; i++)
                        {
                            ziduanming += dtziduan.Rows[i][0] + ",";
                        }
                        //删除结尾字符‘，’
                        ziduanming = ziduanming.TrimEnd(',');

                        sqlcmd = "SELECT " + ziduanming + " FROM  " + dttime.Rows[a]["name"] + "";
                        DataTable dttable = mysql.GetdtTable(sqlcmd);
                        using (SqlBulkCopy bcp = new SqlBulkCopy(scon))
                        {
                            writeR(label1, "" + DateTime.Now.ToString() + "：开始清空远程永久表数据......]");
                            //清空远程永久表
                            scom.CommandText = "TRUNCATE TABLE " + dttime.Rows[a]["name"] + "";
                            scom.ExecuteNonQuery();
                            writeR(label1, "" + DateTime.Now.ToString() + "：清空永久表数据结束。]");
                            bcp.DestinationTableName = dttime.Rows[a]["name"].ToString();
                            writeR(label1, "" + DateTime.Now.ToString() + "：开始推送数据：" + bcp.DestinationTableName + "]");
                            bcp.BulkCopyTimeout = 600;
                            bcp.BatchSize = 1000;
                            for (int i = 0; i < dttable.Columns.Count; i++)
                                bcp.ColumnMappings.Add(dttable.Columns[i].ColumnName.ToString(), dttable.Columns[i].ColumnName.ToString());
                            bcp.WriteToServer(dttable);
                            bcp.Close();
                            writeR(label1, "" + DateTime.Now.ToString() + "：推送结束：" + bcp.DestinationTableName + "]");
                            Debug.WriteLine(DateTime.Now.ToString());
                            Debug.WriteLine(dttime.Rows[a]["name"]);
                        }
                        //writeR(label1, "" + DateTime.Now.ToString() + "：开始更新机型编码---[" + dttime.Rows[a]["name"] + "]");
                        //scom.CommandText = "update A set 机型编码=b.机型编码  from " + dttime.Rows[a]["name"] + " a join 型号表 b on a.机型=b.机型 and a.品牌=b.品牌 and a.品类=b.品类";
                        //scom.ExecuteNonQuery();
                        //writeR(label1, "" + DateTime.Now.ToString() + "：机型编码更新结束---[" + dttime.Rows[a]["name"] + "]");
                    }
                }
            }
            string endtime = DateTime.Now.ToString();
            MessageBox.Show("完成: " + begintime + " --- " + endtime);
            return;
        }
        
        //测试
        private void gettest()
        {
            //1、查找计算列
            string sql = " SELECT NAME FROM sys.computed_columns WHERE object_id = object_id('att彩电属性表') ";
            DataTable dtcolumn = mysql.GetdtTable(sql);
            string columnstr = string.Empty;
            if (dtcolumn.Rows.Count > 0)
            {
                for (int c = 0; c < dtcolumn.Rows.Count; c++)
                {
                    columnstr += "'" + dtcolumn.Rows[c]["name"] + "',";
                }
                columnstr = columnstr.Remove(columnstr.LastIndexOf(","));
            }
            if (columnstr == "")
                columnstr = "''";
            sql= "SELECT COLUMN_NAME FROM  INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='att彩电属性表' AND COLUMN_NAME NOT IN(" + columnstr + ")";
            Debug.WriteLine(columnstr);
        }

        //仅仅推送永久表
        private void tuiyjb()
        {
            string begintime = DateTime.Now.ToString();
            //第二批
            string[] category = new[] { "净水器", "净饮机", "饮水机", "净化器", "电熨斗", "吸尘器", "电风扇", "电暖器", "电磁炉", "电蒸炖锅", "除湿机" };
            //第三批
            category = new[] { "干衣机", "彩电", "冰箱", "冰柜", "洗衣机", "智能机顶盒", "投影仪" };
            //category = new[] { "干衣机", "彩电", "冰箱",  "洗衣机", "智能机顶盒", "投影仪" };
            category = new[] { "智能马桶" };

            category = new[] { "油烟机","燃气灶","消毒柜","热水器" };
            using (SqlConnection scon = new SqlConnection(scon1))
            {
                scon.Open();
                SqlCommand scom = scon.CreateCommand();
                scom.CommandTimeout = 600;
                SqlDataAdapter sda = new SqlDataAdapter(scom);
                //3、后添加永久表数据
                string strcategory = string.Empty;
                for (int i = 0; i < category.Length; i++)
                {
                    strcategory += "'" + category[i] + "',";
                }
                strcategory = strcategory.Remove(strcategory.LastIndexOf(","));

                string sc = "SELECT 'SELECT NAME FROM SYSOBJECTS WHERE TYPE=''U'' AND NAME like '''+组别简称+'_%'+品类+'永久表____年%''  AND NAME NOT LIKE ''%_back'' AND NAME NOT LIKE ''%备份''  and name NOT LIKE ''%2018%''  AND NAME NOT LIKE ''%2017%''  UNION ALL  ' FROM 品类表 WHERE  品类 IN (" + strcategory + ") ORDER BY 品类";//" + MyConfiguration.UserID + "
                //含套餐单品
                sc = "SELECT 'SELECT NAME FROM SYSOBJECTS WHERE TYPE=''U'' AND (NAME like '''+组别简称+'_%线上周度'+品类+'永久表2017年含套餐单品'' or NAME like '''+组别简称+'_%线上周度'+品类+'永久表2016年含套餐单品'')   UNION ALL  ' FROM 品类表 WHERE  品类 IN (" + strcategory + ") ORDER BY 品类";
                DataTable dtSelect = mysql.GetdtTable(sc);
                string strsxt = string.Empty;
                if (dtSelect.Rows.Count > 0)
                {
                    for (int i = 0; i < dtSelect.Rows.Count; i++)
                    {
                        strsxt += "" + dtSelect.Rows[i][0] + "  ";
                    }
                    strsxt = strsxt.Remove(strsxt.ToUpper().LastIndexOf("UNION ALL"));
                    DataTable dttime = mysql.GetdtTable(strsxt);
                    for (int a = 0; a < dttime.Rows.Count; a++)//永久表
                    {
                        writeR(label2, "共 " + (a + 1) + " / " + dttime.Rows.Count + " 个:");
                        //if (a < 132)
                        //    continue;
                        writeR(label1, "" + DateTime.Now.ToString() + "：开始推送永久---[" + dttime.Rows[a]["name"] + "]");
                        //对远程数据库进行操作
                        string sqlcmd = "SELECT COLUMN_NAME FROM  INFORMATION_SCHEMA.COLUMNS where TABLE_NAME='" + dttime.Rows[a]["name"] + "' AND COLUMN_NAME !='ZZID'";
                        DataTable dtziduan = mysql.GetdtTable(sqlcmd);
                        //字段名串
                        string ziduanming = string.Empty;
                        for (int i = 0; i < dtziduan.Rows.Count; i++)
                        {
                            ziduanming += dtziduan.Rows[i][0] + ",";
                        }
                        ziduanming = ziduanming.Replace("销额,", "");//.Replace("机型编码,", "");
                        //删除结尾字符‘，’
                        ziduanming = ziduanming.TrimEnd(',');

                        sqlcmd = "SELECT " + ziduanming + " FROM  " + dttime.Rows[a]["name"] + "";
                        DataTable dttable = mysql.GetdtTable(sqlcmd);
                        using (SqlBulkCopy bcp = new SqlBulkCopy(scon))
                        {
                            writeR(label1, "" + DateTime.Now.ToString() + "：开始清空远程永久表数据......]");
                            //清空远程永久表
                            scom.CommandText = "TRUNCATE TABLE " + dttime.Rows[a]["name"] + "";
                            scom.ExecuteNonQuery();
                            writeR(label1, "" + DateTime.Now.ToString() + "：清空永久表数据结束。]");
                            bcp.DestinationTableName = dttime.Rows[a]["name"].ToString();
                            writeR(label1, "" + DateTime.Now.ToString() + "：开始推送数据："+bcp.DestinationTableName+"]");
                            bcp.BulkCopyTimeout = 600;
                            bcp.BatchSize = 1000;
                            for (int i = 0; i < dttable.Columns.Count; i++)
                                bcp.ColumnMappings.Add(dttable.Columns[i].ColumnName.ToString(), dttable.Columns[i].ColumnName.ToString());
                            bcp.WriteToServer(dttable);
                            bcp.Close();
                            writeR(label1, "" + DateTime.Now.ToString() + "：推送结束：" + bcp.DestinationTableName + "]");
                            Debug.WriteLine(DateTime.Now.ToString());
                            Debug.WriteLine(dttime.Rows[a]["name"]);
                        }
                        writeR(label1, "" + DateTime.Now.ToString() + "：开始更新机型编码---[" + dttime.Rows[a]["name"] + "]");
                        scom.CommandText = "update A set 机型编码=b.机型编码  from " + dttime.Rows[a]["name"] + " a join 型号表 b on a.机型=b.机型 and a.品牌=b.品牌 and a.品类=b.品类";
                        scom.ExecuteNonQuery();
                        writeR(label1, "" + DateTime.Now.ToString() + "：机型编码更新结束---[" + dttime.Rows[a]["name"] + "]");
                    }
                }
            }
            string endtime = DateTime.Now.ToString();
            MessageBox.Show("完成: " + begintime + " --- " + endtime);
            return;
        }

        private void pingIp()
        {
            string assd = "Data Source=124.89.13.18,1433;Initial Catalog=dpcdata;User Id=sa;Password=All_View_Consulting_2014@;";
            string rx = @"\d+.\d+.\d+.\d+";
            string ip = Regex.Match(assd, rx).Value;

            Debug.WriteLine(rx);
            Ping pinsenter = new Ping();
            PingOptions options = new PingOptions();
            options.DontFragment = true;
            string data = "test ping ip";
            byte[] buf = Encoding.ASCII.GetBytes(data);
            //调用同步send 方法发送数据,结果存入reply对象
            PingReply reply = pinsenter.Send("124.89.13.18", 120, buf, options);
            if (reply.Status == IPStatus.Success)
            {
            }
            else
            {
                Debug.WriteLine(reply.Status);
            }
        }

        private void test2()
        {
            string[] specialCagtegory = new string[] { "食品料理机", "料理机", "榨汁机", "食品加工机", "热水器", "电储水热水器", "电即热热水器", "空气能热水器", "燃气热水器", "太阳能热水器" };
            string category = "bx";
            if (specialCagtegory.Contains(category))
            {
                Debug.WriteLine("");
            }
            Debug.WriteLine("");

            string columnName = string.Empty;
            string columnShop = string.Empty; //商场、地市字段

            string tempSqlcmd = "SELECT 字段名称 FROM 数据导出_属性字段设置表 WHERE 品类='冰箱'";
            List<string> exportColumnName = (from r in mysql.GetdtTable(tempSqlcmd).AsEnumerable() select r.Field<string>("字段名称")).ToList<string>();
            if (exportColumnName.Count > 0)
            {
                foreach (string column in exportColumnName)
                    columnName += "," + column;
            }
            else
                columnName = string.Empty;

            tempSqlcmd = "SELECT 字段名称 FROM 数据导出_商场字段设置表 WHERE 品类='冰箱'";
            List<string> exportShopName = (from r in mysql.GetdtTable(tempSqlcmd).AsEnumerable() select r.Field<string>("字段名称")).ToList<string>();
            if (exportShopName.Count > 0)
            {
                foreach (string column in exportShopName)
                    columnShop += "," + column;
            }
            else
                columnShop = string.Empty;

            tempSqlcmd = "SELECT 输出字段名 FROM 数据导出_分段设置表 WHERE 品类='冰箱' GROUP BY 输出字段名 ORDER BY 输出字段名";
            List<string> exportSplitValueName = (from r in mysql.GetdtTable(tempSqlcmd).AsEnumerable() select r.Field<string>("输出字段名")).ToList<string>();
            string exportString = string.Empty;

            #region 分段设置表 中存在分段规则情况
            if (exportSplitValueName.Count != 0)
            {
                foreach (string segtion in exportSplitValueName)
                {
                    tempSqlcmd = "SELECT 输出字段 " + segtion + ",ISNULL(属性名1,'') 属性名1,属性1,分段字段名,段起始值,段终止值 " +
                        "FROM 数据导出_分段设置表 WHERE 品类='冰箱' AND 输出字段名='" + segtion + "'";
                    DataTable segtionInformation = mysql.GetdtTable(tempSqlcmd);
                    if (segtionInformation.Rows.Count == 0)
                        continue;

                    #region wgh 判断是否有新的属性分段（新方法）测试 （单个属性）
                    #region 存在一个属性段
                    if (segtionInformation.Rows[0]["属性名1"].ToString() != "")
                    {
                        string SXValue = string.Empty, SXValueCount = string.Empty;
                        var queSXName = from p in segtionInformation.AsEnumerable()//属性名1
                                        group p by p.Field<string>("属性名1") into g
                                        select g;
                        foreach (var sxname in queSXName)
                        {
                            var queSXValue = from p in segtionInformation.AsEnumerable()//属性1
                                             where p.Field<string>("属性名1").ToUpper().Trim() == sxname.Key.ToString().ToUpper().Trim()
                                             group p by p.Field<string>("属性1") into g
                                             select g;
                            foreach (var sxvalue in queSXValue)//属性值+段
                            {
                                SXValue = " WHEN " + sxname.Key.ToString() + "='" + sxvalue.Key.ToString() + "'  THEN CASE ";
                                var queSegment = from p in segtionInformation.AsEnumerable()
                                                 where p.Field<string>("属性名1").ToUpper().Trim() == sxname.Key.ToString().ToUpper().Trim() && p.Field<string>("属性1").ToUpper().Trim() == sxvalue.Key.ToString().ToUpper().Trim()
                                                 group p by new { sx1 = p.Field<decimal>("段起始值"), sx2 = p.Field<decimal>("段终止值"), fdzdm = p.Field<string>("分段字段名"), sczd = p.Field<string>(segtion) } into g
                                                 select g;
                                string SXV = string.Empty;
                                foreach (var qsegment in queSegment)//属性值+段
                                    SXV += " WHEN " + qsegment.Key.fdzdm.ToUpper().Trim() + ">=" + qsegment.Key.sx1 + " AND " + qsegment.Key.fdzdm.ToUpper().Trim() + "<" + qsegment.Key.sx2 + " THEN '" + qsegment.Key.sczd + "' ";
                                SXValueCount += SXValue + (SXV + " END ") + "";
                            }
                            exportString += ",( CASE " + SXValueCount + " END ) AS " + segtion + "";
                        }
                    }
                    #endregion

                    #region 不存在属性段
                    else
                    {
                        string SXValue = string.Empty, SXValueCount = string.Empty;
                        var queSegment = from p in segtionInformation.AsEnumerable()
                                         group p by new { sx1 = p.Field<decimal>("段起始值"), sx2 = p.Field<decimal>("段终止值"), fdzdm = p.Field<string>("分段字段名"), sczd = p.Field<string>(segtion) } into g
                                         select g;
                        string SXV = string.Empty;
                        foreach (var qsegment in queSegment)//属性值+段
                            SXV += " WHEN " + qsegment.Key.fdzdm.ToUpper().Trim() + ">=" + qsegment.Key.sx1 + " AND " + qsegment.Key.fdzdm.ToUpper().Trim() + "<" + qsegment.Key.sx2 + " THEN '" + qsegment.Key.sczd + "' ";
                        SXValueCount += SXValue + (SXV + " END ") + "";
                        exportString += ",( CASE " + SXValueCount + " ) AS " + segtion + "";
                    }
                    #endregion

                    #endregion
                }
            }
            else
            { }
            Debug.WriteLine("");
            #endregion

            #region 线下厨电套餐周度同比库今比昔
            //permenentCountString = "SELECT SUM(销量)销量 FROM " +
            //"(SELECT 商场店名,机型编码,周度,单价,SUM(销量) 销量 FROM KA_线下周度厨电套餐永久表20" + combinedConditions[i].Substring(0, 2) + "年" + danpin + " Y " +
            //"WHERE Y.周度='" + combinedConditions[i] + "' AND EXISTS(SELECT L.商场店名 FROM KA_线下周度厨电套餐永久表20" + (Int32.Parse(combinedConditions[i].Substring(0, 2)) - 1) + "年" + danpin + "  L " +
            //"WHERE L.商场店名=Y.商场店名 AND L.周度='" + (Int32.Parse(combinedConditions[i].Substring(0, 2)) - 1) + combinedConditions[i].Substring(2, 3) + "')  GROUP BY 商场店名,机型编码,周度,单价) T";
            string permenentCountString = "SELECT COUNT(*) FROM ( " +
                "SELECT 商场店名,机型编码,周度,单价,SUM(销量) 销量 FROM IW_线下周度冰箱永久表2017年 A WHERE 周度='17w50' AND EXISTS(SELECT 商场店名 FROM IW_线下周度冰箱永久表2016年 B WHERE 周度='16w50'  AND  A.商场店名=B.商场店名) GROUP BY 商场店名,机型编码,周度,单价)YJB ";
           //DataTable permanentCount = mysql.GetdtTable(permenentCountString);
       
            exportString = "SELECT 销量,单价*销量 AS 销额,周度, SXB.机型编码,SXB.品牌" + columnName + columnShop + exportString + " FROM (" +
                "SELECT 商场店名,机型编码,周度,单价,SUM(销量) 销量 FROM IW_线下周度冰箱永久表2017年 A WHERE 周度='17w50' AND EXISTS(SELECT 商场店名 FROM IW_线下周度冰箱永久表2016年 B WHERE 周度='16w50' AND EXISTS (SELECT 机型编码 FROM ATT冰箱属性表 C WHERE A.机型编码=C.机型编码) AND  A.商场店名=B.商场店名) GROUP BY 商场店名,机型编码,周度,单价)YJB  " +
                "JOIN " +
                "(SELECT SA.机型编码,SA.品牌,SA.机型" + columnName + "  FROM ATT冰箱属性表 SA JOIN 型号表 XHB ON SA.机型编码=XHB.机型编码)SXB ON YJB.机型编码=SXB.机型编码 " +
                "JOIN (SELECT 品牌类型,品牌 FROM 品牌表) SP ON SXB.品牌=SP.品牌 " +
                "JOIN (SELECT 商场编码"+columnShop+" FROM 线下商场库 SC JOIN 县市区划表 XS ON 地市=地市名称 AND 省份=省份名称 AND 县市=县市名称)SCK ON 商场店名=商场编码";
           string  linkedCountString = "SELECT COUNT(*) FROM (" + exportString + ") T";

           Debug.WriteLine("");

            #endregion
        }

        private void test()
        {
            string sql = "SELECT DISTINCT 输出字段名,columnname='分段表' FROM 数据导出_分段设置表 WHERE 品类='干衣机' "+
                "UNION ALL "+
                "SELECT DISTINCT 字段名称,columnname='属性表' FROM 数据导出_属性字段设置表 WHERE 品类='榨汁机' "+
                "UNION ALL "+
                "SELECT DISTINCT 字段名称,columnname='商场表' FROM 数据导出_商场字段设置表 WHERE 品类='榨汁机' AND 字段名称!='连锁名称'";
            DataTable dtx = mysql.GetdtTable(sql);
            BindTreeViewFatherNodeData(dtx, "columnname");
        }
        private void BindTreeViewFatherNodeData(DataTable dtTreeData, string fatherColumnName)
        {
            if (dtTreeData.Rows.Count > 0)
            {
                var columnName = from p in dtTreeData.AsEnumerable()
                                 group p by new { column = p.Field<object>(fatherColumnName) } into g
                                 select g.Key;
                if (columnName.Count() > 0)
                {
                    foreach (var colu in columnName)
                    {
                        Debug.WriteLine(colu);
                    }
                }
            }
        }


        //修改周度-1
        private void yjb()
        {
            string begintime = DateTime.Now.ToString();
            //1、先修改周度 -1
            string sc = "select 'update '+name+' set 周度=(case when convert(int,substring(周度,4,2))<=10 then (substring(周度,1,3)+''0'')+convert(nvarchar(10),(substring(周度,4,2)-1)) else substring(周度,1,3)+''''+convert(nvarchar(10),(substring(周度,4,2)-1)) end)' from sysobjects where type='u' and (name like '%周度%永久表2017%' or name like '%周度%永久表2016%') AND NAME NOT LIKE '%_BACKUP' and name not like '%bf%' and name not like '%三星%'";
            DataTable dtSelect = mysql.GetdtTable(sc);
            string strsxt = string.Empty;
            if (dtSelect.Rows.Count > 0)
            {
                writeR(label2, "共 " + 0 + " / " + dtSelect.Rows.Count + " 个:");
                for (int i = 0; i < dtSelect.Rows.Count; i++)
                {
                    if (i < 215)
                        continue;
                    //彩电 线下周度
                    writeR(label2, "开始 " + (i + 1) + " / " + dtSelect.Rows.Count + " 个。。。");
                    try
                    {
                        mysql.ExecuteNonQuery(dtSelect.Rows[i][0].ToString());
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine(ex.Message);
                        MessageBox.Show(ex.Message);
                    }
                }
            }
            //2、把 W00的数据 移动至去年永久表
            //①、线上周度
            //--1
            sc = "select  'INSERT INTO '+replace(name,'2017','2016')+'(月度,周度,品类,品牌,机型编码,机型,单价,销量,电商,评论人,评论日期,总体评价,商品名称,页面信息,促销信息,旗舰店,dataflag,销售类型,货仓,月销量,COLOUR,SIZE) "+
                "SELECT 月度,周度=''16W52'',品类,品牌,机型编码,机型,单价,销量,电商,评论人,评论日期,总体评价,商品名称,页面信息,促销信息,旗舰店,dataflag,销售类型,货仓,月销量,COLOUR,SIZE FROM '+name+' WHERE 周度=''17W00''' from sysobjects where type='u' and (name like '%线上周度%永久表2017%') and name not like '%三星%' AND NAME NOT LIKE '%_BACKUP' and name not like '%bf%'";
            DataTable dtSelectin = mysql.GetdtTable(sc);
            if (dtSelectin.Rows.Count > 0)
            {
                writeR(label2, "线上开始处理 17W00  " + 0 + " / " + dtSelectin.Rows.Count + " 个:");

                for (int i = 0; i < dtSelectin.Rows.Count; i++)
                {
                    writeR(label2, "线上开始处理 17W00   " + (i + 1) + " / " + dtSelectin.Rows.Count + " 个。。。");
                    try
                    {
                        mysql.ExecuteNonQuery(dtSelectin.Rows[i][0].ToString());
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine(ex.Message);
                        MessageBox.Show(ex.Message);
                    }
                }
            }
            //2、
            sc = "select  'INSERT INTO '+replace(name,'2016','2015')+'(月度,周度,品类,品牌,机型编码,机型,单价,销量,电商,评论人,评论日期,总体评价,商品名称,页面信息,促销信息,旗舰店,dataflag,销售类型,货仓,月销量) "+
                "SELECT 月度,周度=''15W53'',品类,品牌,机型编码,机型,单价,销量,电商,评论人,评论日期,总体评价,商品名称,页面信息,促销信息,旗舰店,dataflag,销售类型,货仓,月销量 FROM '+name+' WHERE 周度=''16W00''' from sysobjects where type='u' and (name like '%线上周度%永久表2016%') and name not like '%三星%' AND NAME NOT LIKE '%_BACKUP' and name not like '%bf%'";
             dtSelectin = mysql.GetdtTable(sc);
             if (dtSelectin.Rows.Count > 0)
             {
                 writeR(label2, "线上开始处理 16W00  " + 0 + " / " + dtSelectin.Rows.Count + " 个:");

                 for (int i = 0; i < dtSelectin.Rows.Count; i++)
                 {
                     writeR(label2, "线上开始处理 16W00   " + (i + 1) + " / " + dtSelectin.Rows.Count + " 个。。。");
                     try
                     {
                         mysql.ExecuteNonQuery(dtSelectin.Rows[i][0].ToString());
                     }
                     catch (Exception ex)
                     {
                         Debug.WriteLine(ex.Message);
                         MessageBox.Show(ex.Message);
                     }
                 }
             }
            //②、线下周度 
            //--1
             sc = "select 'INSERT INTO '+replace(name,'2017','2016')+'(周度,商场店名,商品名称,品类,品牌,机型编码,机型,单价,销量,数据来源,dataflag,商场子码) "+
                 " SELECT 周度=''16W52'',商场店名,商品名称,品类,品牌,机型编码,机型,单价,销量,数据来源,dataflag,商场子码 FROM '+name+' WHERE 周度=''17W00''' from sysobjects where type='u' and (name like '%线下周度%永久表2017%') and name not like '%三星%' AND NAME NOT LIKE '%_BACKUP' and name not like '%bf%'";
             dtSelectin = mysql.GetdtTable(sc);
             if (dtSelectin.Rows.Count > 0)
             {
                 writeR(label2, "线下开始处理 17W00  " + 0 + " / " + dtSelectin.Rows.Count + " 个:");

                 for (int i = 0; i < dtSelectin.Rows.Count; i++)
                 {
                     writeR(label2, "线下开始处理 17W00   " + (i + 1) + " / " + dtSelectin.Rows.Count + " 个。。。");
                     try
                     {
                         mysql.ExecuteNonQuery(dtSelectin.Rows[i][0].ToString());
                     }
                     catch (Exception ex)
                     {
                         Debug.WriteLine(ex.Message);
                         MessageBox.Show(ex.Message);
                     }
                 }
             }
             //--2
             sc = "select 'INSERT INTO '+replace(name,'2016','2015')+'(周度,商场店名,商品名称,品类,品牌,机型编码,机型,单价,销量,数据来源,dataflag,商场子码) "+
                 "   SELECT 周度=''15W53'',商场店名,商品名称,品类,品牌,机型编码,机型,单价,销量,数据来源,dataflag,商场子码 FROM '+name+' WHERE 周度=''16W00''' from sysobjects where type='u' and (name like '%线下周度%永久表2016%') and name not like '%三星%' AND NAME NOT LIKE '%_BACKUP' and name not like '%bf%'";
             dtSelectin = mysql.GetdtTable(sc);
             if (dtSelectin.Rows.Count > 0)
             {
                 writeR(label2, "线下开始处理 16W00  " + 0 + " / " + dtSelectin.Rows.Count + " 个:");

                 for (int i = 0; i < dtSelectin.Rows.Count; i++)
                 {
                     writeR(label2, "线下开始处理 16W00   " + (i + 1) + " / " + dtSelectin.Rows.Count + " 个。。。");
                     try
                     {
                         mysql.ExecuteNonQuery(dtSelectin.Rows[i][0].ToString());
                     }
                     catch (Exception ex)
                     {
                         Debug.WriteLine(ex.Message);
                         MessageBox.Show(ex.Message);
                     }
                 }
             }
            //--3、删除 W00数据
             //①、线上、线下周度
             sc = "select 'delete '+name+' WHERE 周度=''17W00''' from sysobjects where type='u' and (name like '%周度%永久表2017%' or name like '%周度%永久表2017%') and name not like '%三星%' AND NAME NOT LIKE '%_BACKUP' and name not like '%bf%'";
             dtSelectin = mysql.GetdtTable(sc);
             if (dtSelectin.Rows.Count > 0)
             {
                 writeR(label2, "线下\\线上开始处理删除 17W00  " + 0 + " / " + dtSelectin.Rows.Count + " 个:");

                 for (int i = 0; i < dtSelectin.Rows.Count; i++)
                 {
                     writeR(label2, "线下\\线上开始处理删除 17W00  " + (i + 1) + " / " + dtSelectin.Rows.Count + " 个。。。");
                     try
                     {
                         mysql.ExecuteNonQuery(dtSelectin.Rows[i][0].ToString());
                     }
                     catch (Exception ex)
                     {
                         Debug.WriteLine(ex.Message);
                         MessageBox.Show(ex.Message);
                     }
                 }
             }
            //2
             sc = "select 'delete '+name+' WHERE 周度=''16W00''' from sysobjects where type='u' and (name like '%周度%永久表2016%' or name like '%周度%永久表2016%') and name not like '%三星%' AND NAME NOT LIKE '%_BACKUP' and name not like '%bf%'";
             dtSelectin = mysql.GetdtTable(sc);
             if (dtSelectin.Rows.Count > 0)
             {
                 writeR(label2, "线下\\线上开始处理删除 16W00  " + 0 + " / " + dtSelectin.Rows.Count + " 个:");

                 for (int i = 0; i < dtSelectin.Rows.Count; i++)
                 {
                     writeR(label2, "线下\\线上开始处理删除 16W00  " + (i + 1) + " / " + dtSelectin.Rows.Count + " 个。。。");
                     try
                     {
                         mysql.ExecuteNonQuery(dtSelectin.Rows[i][0].ToString());
                     }
                     catch (Exception ex)
                     {
                         Debug.WriteLine(ex.Message);
                         MessageBox.Show(ex.Message);
                     }
                 }
             }

            string endtime = DateTime.Now.ToString();
            MessageBox.Show("完成: " + begintime + " --- " + endtime);
            this.button1.Enabled = true;
            return;
        }
    }
}
