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
using System.IO;

namespace AVC_ClareData
{
    public partial class InsertXian_sjxh : Form
    {
        //string scon1 = "Data Source=124.89.13.18,1433;Initial Catalog=dpcdata;User Id=sa;Password=All_View_Consulting_2014@;";
        string scon1 = string.Empty;//"Data Source=192.168.2.236;Initial Catalog=dpcdata;User Id=sa;Password=All_View_Consulting_2014@;";

        MySqlConnection mysql = new MySqlConnection();
        /// <summary>
        /// 异常信息文件相对路径\存放路径
        /// </summary>
        string errorfilepath = AppDomain.CurrentDomain.BaseDirectory + "\\errorlog\\";//路径
        /// <summary>
        /// 设置时间(每天开始时间) 默认早晨6点开始
        /// </summary>
        private int setTime = 6;

        public InsertXian_sjxh()
        {
            InitializeComponent();
            this.progressBar1.Visible = false;
            //scon1 = System.Configuration.ConfigurationManager.ConnectionStrings["strSql"].ConnectionString.ToString();//获取配置文件ip、密码ect
        }

        private void button1_Click(object sender, EventArgs e)
        {
            begin = false;
            new Thread(new ThreadStart(delegate
            {
                richtextwriteR(richTextBox1, "");
                beginWork();
            })).Start();
        }

        private void writeR(Label c, string aa)
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

        private void textwriteR(TextBox c, string aa)
        {
            try
            {
                lock (this)
                {
                    c.Invoke(new ThreadStart(delegate()
                    {
                        if (aa == "" || aa == null)
                            c.Text = "";
                        else
                            c.Text +="\n\r"+ aa + "\n\r";
                    }));
                }
            }
            catch { }
        }

        /// <summary>
        /// RichTextBox 控件线程调用辅助公用方法
        /// </summary>
        /// <param name="c">RichTextBox name</param>
        /// <param name="strtext">赋值的文本内容</param>
        private void richtextwriteR(RichTextBox c, string strtext)
        {
            try
            {
                lock (this)
                {
                    c.Invoke(new ThreadStart(delegate()
                    {
                        if (strtext == "" || strtext == null)
                            c.Text = "";
                        c.AppendText(strtext + "\r\n");
                        c.Focus();
                    }));
                }
            }
            catch { }
        }

        int ix = 0; bool begin = true;
        private void pTimer_Elapsed(object sender, System.Timers.ElapsedEventArgs e)
        {
            while (!this.IsHandleCreated)
            {
                ;
            }
            this.Invoke(new ThreadStart(delegate()
            {
                if (textBox_setTime.Text != "")
                    setTime = Convert.ToInt32(textBox_setTime.Text.Trim());
                else
                    setTime = 5;
            }));
            ix += 1;
            writeR(label_time, "当前时间：" + DateTime.Now.ToString());
            if (begin == true)
            {
                if (Convert.ToInt32(DateTime.Now.Hour) == setTime && Convert.ToInt32(DateTime.Now.Minute) == 1)//每天的凌晨6点执行
                {
                    richtextwriteR(richTextBox1, "");
                    ix = 0;
                    begin = false;
                    Debug.WriteLine("开始执行：" + DateTime.Now.ToString());
                    richtextwriteR(richTextBox1, DateTime.Now.ToString() + "：开始执行......");
                    button1_Click(null, null);
                }
            }
        }

        private void beginWork()
        {
            string begintime = DateTime.Now.ToString();
            //查找品类
            string strSql = "SELECT DISTINCT 品类 FROM 品类表 WHERE 组别简称!='' AND 品类 NOT IN('智能手机') ";
            DataTable dtCategory = new DataTable();
            bool sfx = true;
            int ljcount = 1;

            #region 判断
            while (sfx)
            {
                try
                {
                    dtCategory = mysql.GetdtTable(strSql);
                    sfx = false;
                    writeR(this.label_times, "");
                }
                catch (SqlException ex)
                {
                    if (!Directory.Exists(errorfilepath))
                        Directory.CreateDirectory(errorfilepath);
                    if (ex.Number == 53)
                    {
                        writeR(this.label_times, "sqlserver数据库连接失败,正在重新连接--[" + ljcount + "]........");
                        richtextwriteR(richTextBox1, DateTime.Now.ToString() + "：sqlserver数据库连接失败,正在重新连接--[" + ljcount + "]........");
                        Thread.Sleep(1000 * 60);//断网重新链接必须要等待2s以上
                    }
                    else
                    {
                        writeR(this.label_times, "sqlserver错误信息：" + ex.Message + "");
                        richtextwriteR(richTextBox1, DateTime.Now.ToString() + "：sqlserver错误信息：" + ex.Message + "");
                        Thread.Sleep(2000);//断网重新链接必须要等待2s以上
                    }
                    if (!Directory.Exists(errorfilepath))
                        Directory.CreateDirectory(errorfilepath);
                    File.AppendAllText(errorfilepath + "sqlserver错误信息" + DateTime.Now.ToString("yyyy-MM-dd") + "!.log", DateTime.Now.ToString() + " ------ " + ex.Message + "\r\n----------------------------------------\r\n");
                    ljcount += 1;
                }
            }
            #endregion

            writeR(this.label_times, "");

            for (int a = 0; a < dtCategory.Rows.Count; a++)//永久表
            {
                writeR(label2, "共 " + (a + 1) + " / " + dtCategory.Rows.Count + " 个:");
                writeR(label1, "" + DateTime.Now.ToString() + "：开始删除商家型号表数据......]");
                strSql = "DELETE FROM 商家型号对照表 WHERE 品类='" + dtCategory.Rows[a]["品类"] + "'";
                #region 判断
                sfx = true;
                while (sfx)
                {
                    try
                    {
                        mysql.ExecuteNonQuery(strSql);
                        sfx = false;
                        writeR(this.label_times, "");
                    }
                    catch (SqlException ex)
                    {
                        if (!Directory.Exists(errorfilepath))
                            Directory.CreateDirectory(errorfilepath);
                        if (ex.Number == 53)
                        {
                            Thread.Sleep(1000 * 60);//断网重新链接必须要等待2s以上
                            writeR(this.label_times, "sqlserver数据库连接失败,正在重新连接--[" + ljcount + "]........");
                        }
                        else
                        {
                            writeR(this.label_times, "sqlserver错误信息：" + ex.Message + "");
                            Thread.Sleep(2000);//断网重新链接必须要等待2s以上
                        }
                        if (!Directory.Exists(errorfilepath))
                            Directory.CreateDirectory(errorfilepath);
                        File.AppendAllText(errorfilepath + "sqlserver错误信息" + DateTime.Now.ToString("yyyy-MM-dd") + "!.log", DateTime.Now.ToString() + " ------ " + ex.Message + "\r\n----------------------------------------\r\n");
                        ljcount += 1;
                    }
                }
                #endregion
                writeR(this.label_times, "");
                writeR(label1, "" + DateTime.Now.ToString() + "：删除商家型号表数据结束。");
                richtextwriteR(richTextBox1, DateTime.Now.ToString() + "：删除商家型号表数据结束----[" + dtCategory.Rows[a]["品类"] + "]。");

                strSql = "INSERT INTO 商家型号对照表(机型,商家机型,品类,品牌,写入日期)  SELECT 机型,商家机型,品类,品牌,写入日期 FROM OPENROWSET('SQLOLEDB','124.89.13.18,1433';'sa';'All_View_Consulting_2014@',DPCDATA.DBO.商家型号对照表)WHERE 品类='" + dtCategory.Rows[a]["品类"] + "'";
                writeR(label1, "" + DateTime.Now.ToString() + "：开始抽取---[" + dtCategory.Rows[a]["品类"] + "]");
                richtextwriteR(richTextBox1, DateTime.Now.ToString() + "：开始抽取---[" + dtCategory.Rows[a]["品类"] + "]");
                //查询远程服务器商家型号表
                #region 判断
                sfx = true;
                while (sfx)
                {
                    try
                    {
                        mysql.ExecuteNonQuery(strSql);
                        sfx = false;
                        writeR(this.label_times, "");
                    }
                    catch (SqlException ex)
                    {
                        if (!Directory.Exists(errorfilepath))
                            Directory.CreateDirectory(errorfilepath);
                        if (ex.Number == 53)
                        {
                            Thread.Sleep(1000 * 60);//断网重新链接必须要等待2s以上
                            writeR(this.label_times, "删除sqlserver数据库连接失败,正在重新连接--[" + ljcount + "]........");
                        }
                        else
                        {
                            writeR(this.label_times, "删除sqlserver错误信息：" + ex.Message + "");
                            Thread.Sleep(2000);//断网重新链接必须要等待2s以上
                        }
                        if (!Directory.Exists(errorfilepath))
                            Directory.CreateDirectory(errorfilepath);
                        File.AppendAllText(errorfilepath + "删除sqlserver错误信息" + DateTime.Now.ToString("yyyy-MM-dd") + "!.log", DateTime.Now.ToString() + " ------ " + ex.Message + "\r\n----------------------------------------\r\n");
                        ljcount += 1;
                    }
                }

                #endregion
                writeR(this.label_times, "");
                writeR(label1, "" + DateTime.Now.ToString() + "：推送结束：" + dtCategory.Rows[a]["品类"] + "]");
                richtextwriteR(richTextBox1, DateTime.Now.ToString() + "：推送结束：" + dtCategory.Rows[a]["品类"] + "]");
            }
            //抽取完成!!!!
            writeR(label1, "" + DateTime.Now.ToString() + "：抽取完成!!!");
            richtextwriteR(richTextBox1, DateTime.Now.ToString() + "：抽取完成!!!");
            //查询本地验证
            strSql = "SELECT COUNT(DISTINCT 品类) FROM 商家型号对照表";
            DataTable dtCount = mysql.GetdtTable(strSql);
            if (dtCount.Rows.Count > 0)
            {
                richtextwriteR(richTextBox1, DateTime.Now.ToString() + "：本地共  " + dtCount.Rows[0][0] + " 个品类");
            }
            else
                richtextwriteR(richTextBox1, DateTime.Now.ToString() + "：本地共  0 个品类,请检查是否出现问题!!!!");

            TimeSpan ts = DateTime.Now - Convert.ToDateTime(begintime);
            richtextwriteR(richTextBox1, DateTime.Now.ToString() + "：共耗时：" + ts);
            begin = true;
        }

        private void InsertXian_sjxh_Load(object sender, EventArgs e)
        {
            System.Timers.Timer pTimer = new System.Timers.Timer(1000);//每隔5s执行一次
            pTimer.Elapsed += pTimer_Elapsed;//委托
            pTimer.AutoReset = true;//获取定时器自动执行
            pTimer.Enabled = true;
            Control.CheckForIllegalCrossThreadCalls = false;//调用线程后台调用，不会影响控件的显示
        }

        private void InsertXian_sjxh_FormClosed(object sender, FormClosedEventArgs e)
        {
            Process.GetCurrentProcess().Kill();
        }
    }
}
