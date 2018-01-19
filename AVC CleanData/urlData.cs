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
using AVC_ClareData.PublicClass;

namespace AVC_ClareData
{
    public partial class urlData : Form
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

        public urlData()
        {
            InitializeComponent();
        }

        /// <summary>
        /// lable 控件线程调用辅助公用方法
        /// </summary>
        /// <param name="c">lable name</param>
        /// <param name="strtext">赋值的文本内容</param>
        private void writeR(Label c, string strtext)
        {
            try
            {
                lock (this)
                {
                    c.Invoke(new ThreadStart(delegate()
                    {
                        c.Text = strtext;
                    }));
                }
            }
            catch { }
        }

        /// <summary>
        ///  TextBox 控件线程调用辅助公用方法
        /// </summary>
        /// <param name="c">TextBox name</param>
        /// <param name="strtext">赋值的文本内容</param>
        private void textwriteR(TextBox c, string strtext)
        {
            try
            {
                lock (this)
                {
                    c.Invoke(new ThreadStart(delegate()
                    {
                        if (strtext == "" || strtext == null)
                            c.Text = "";
                        else
                            c.Text += "\n\r" + strtext + "\n\r";
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
            DateTime dttime = DateTime.Now.AddDays(-7);//查询上周
            string week = DateTimeOperation.Week(dttime);
            writeR(label_time, "当前时间：" + DateTime.Now.ToString());
            writeR(label_week, "当前周度： ( " + week + " )");
            if (begin == true)
            {
                if (Convert.ToInt32(DateTime.Now.Hour) == setTime && Convert.ToInt32(DateTime.Now.Minute) == 1)//每天的凌晨6点执行
                {
                    richtextwriteR(richTextBox_yjbUrl, "");
                    ix = 0;
                    begin = false;
                    Debug.WriteLine("开始执行：" + DateTime.Now.ToString());
                    richtextwriteR(richTextBox_yjbUrl, DateTime.Now.ToString() + "：开始执行......");

                    button_begin_Click(null, null);
                }
            }
        }

        private void urlData_Load(object sender, EventArgs e)
        {
            System.Timers.Timer pTimer = new System.Timers.Timer(1000);//每隔5s执行一次
            pTimer.Elapsed += pTimer_Elapsed;//委托
            pTimer.AutoReset = true;//获取定时器自动执行
            pTimer.Enabled = true;
            Control.CheckForIllegalCrossThreadCalls = false;//调用线程后台调用，不会影响控件的显示
        }

        private void urlData_FormClosed(object sender, FormClosedEventArgs e)
        {
            Process.GetCurrentProcess().Kill();
        }
        //永久表url,查询 18年永久表,如果19年，需要修改
        private void workYJB()
        {
            DateTime dttime = DateTime.Now.AddDays(-7);//查询上周
            string week = DateTimeOperation.Week(dttime);
            //查询18年永久表
            string strSql = "SELECT 'INSERT INTO CHDATA.DBO.URLDATA (品类,品牌,机型,页面信息,旗舰店,写入日期,电商,need)  SELECT A.品类,A.品牌,A.机型,A.页面信息,A.旗舰店,写入日期=getdate(),A.电商,need=1 FROM (SELECT 品类,品牌,max(机型)机型,页面信息,max(旗舰店)旗舰店,电商 FROM '+NAME+' WHERE 周度=''" + week + "'' AND 电商!=''淘宝网'' AND 页面信息 LIKE ''http%'' AND 页面信息 NOT LIKE ''%chaoshi%'' AND 页面信息 NOT LIKE ''%已下架%'' GROUP BY 品牌,品类,页面信息,电商)A LEFT JOIN (SELECT 页面信息,电商 FROM CHDATA.DBO.URLDATA WHERE need=1)B ON A.页面信息=B.页面信息 WHERE B.页面信息 IS NULL' FROM SYSOBJECTS WHERE TYPE='U' AND NAME LIKE '%线上周度%永久表2018年' AND NAME NOT LIKE '%三星%' ";
            DataTable dtYjbTable = new DataTable();
            bool sfx = true;
            int ljcount = 1;

            #region 判断

            while (sfx)
            {
                try
                {
                    richtextwriteR(richTextBox_yjbUrl, DateTime.Now.ToString() + "：正在查询永久表......");
                    dtYjbTable = mysql.GetdtTable(strSql);
                    richtextwriteR(richTextBox_yjbUrl, DateTime.Now.ToString() + "：查询永久表结束。");
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
                        richtextwriteR(richTextBox_yjbUrl, DateTime.Now.ToString() + "：sqlserver数据库连接失败,正在重新连接--[" + ljcount + "]........");
                        Thread.Sleep(1000 * 60);//断网重新链接必须要等待2s以上
                    }
                    else
                    {
                        writeR(this.label_times, "sqlserver错误信息：" + ex.Message + "");
                        richtextwriteR(richTextBox_yjbUrl, DateTime.Now.ToString() + "：sqlserver错误信息：" + ex.Message + "");
                        Thread.Sleep(2000);//断网重新链接必须要等待2s以上
                    }
                    if (!Directory.Exists(errorfilepath))
                        Directory.CreateDirectory(errorfilepath);
                    File.AppendAllText(errorfilepath + "永久表推送url错误信息" + DateTime.Now.ToString("yyyy-MM-dd") + "!.log", DateTime.Now.ToString() + " ------ " + ex.Message + "\r\n----------------------------------------\r\n");
                    ljcount += 1;
                }
            }

            #endregion

            if (dtYjbTable.Rows.Count > 0)
            {
                richtextwriteR(richTextBox_yjbUrl, DateTime.Now.ToString() + "：共  " + dtYjbTable.Rows.Count + "  个永久表");
                for (int i = 0; i < dtYjbTable.Rows.Count; i++)
                {

                    strSql = dtYjbTable.Rows[i][0].ToString();

                    //正则提取 品类
                    string StrRegex = @"(?<=线上周度).*(?=永久表\d+年)";
                    string ValueRegex = Regex.Match(strSql, StrRegex).ToString();

                    #region 判断

                    sfx = true;
                    ljcount = 1;
                    while (sfx)
                    {
                        try
                        {
                            //int index = mysql.ExecuteNonQuery(strSql);
                            int index = 0;
                            richtextwriteR(richTextBox_yjbUrl, DateTime.Now.ToString() + "：开始 " + (i + 1) + " / 共  " + dtYjbTable.Rows.Count + "  个");
                            richtextwriteR(richTextBox_yjbUrl, "品类：" + ValueRegex + ",影响：" + index + " 条");
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
                                richtextwriteR(richTextBox_yjbUrl, DateTime.Now.ToString() + "：sqlserver数据库连接失败,正在重新连接--[" + ljcount + "]........");
                                Thread.Sleep(1000 * 60);//断网重新链接必须要等待2s以上
                            }
                            else
                            {
                                writeR(this.label_times, "sqlserver错误信息：" + ex.Message + "");
                                richtextwriteR(richTextBox_yjbUrl, DateTime.Now.ToString() + "：sqlserver错误信息：" + ex.Message + "");
                                Thread.Sleep(2000);//断网重新链接必须要等待2s以上
                            }
                            if (!Directory.Exists(errorfilepath))
                                Directory.CreateDirectory(errorfilepath);
                            File.AppendAllText(errorfilepath + "永久表推送url错误信息" + DateTime.Now.ToString("yyyy-MM-dd") + "!.log", DateTime.Now.ToString() + " ------ " + ex.Message + "\r\n----------------------------------------\r\n");
                            ljcount += 1;
                        }
                    }

                    #endregion
                }
            }
            begin = true;
        }

        private void button_begin_Click(object sender, EventArgs e)
        {
            begin = false;
            new Thread(new ThreadStart(delegate
            {
                richtextwriteR(richTextBox_yjbUrl, "");

                workYJB();
            })).Start();
        }
        
    }
}
