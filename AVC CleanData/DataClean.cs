using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;
using System.Diagnostics;
using System.IO;
using System.Threading;
using System.Runtime.InteropServices;
using AVC_ClareData.PublicClass;
using AVC_ClareData.Model;
using System.Text.RegularExpressions;
using System.Reflection;
using System.Runtime.Serialization.Formatters.Binary;
using System.Security.Cryptography;

namespace AVC_ClareData
{
    public partial class DataClean : Form
    {
        //定义句柄变量
        public static IntPtr hwnd;
        string message = string.Empty;
        //定义进程ID变量
        public static int pid = 0;
        UserSetModel userModel = new UserSetModel();
        MySqlConnection mySql = new MySqlConnection();
        //获取进程文件id
        [DllImport("User32.dll", CharSet = CharSet.Auto)]
        public static extern int GetWindowThreadProcessId(IntPtr hwnd, out int ID);
        string path = string.Empty;
        //使用句柄隐藏窗口
        [DllImport("user32.dll", EntryPoint = "ShowWindow", SetLastError = true)]
        static extern bool ShowWindow(IntPtr hWnd, uint nCmdShow);//nCmdShow ：=0 ，窗体不显示，=1,显示
        /// <summary>
        /// 特殊品牌处理
        /// </summary>
        string[] specialBrands = { "Adidas/阿迪达斯", "ANTA/安踏", "Deerway/德尔惠", "erke/鸿星尔克", "Kappa/背靠背", "Lining/李宁", "nike Air Jordan/乔丹", "Nike/耐克", "Puma/彪马", "乔丹" };
        /// <summary>
        ///   只保留 "-" 之前的字符串
        /// </summary>
        string[] specialBrands2 = { "ANTA/安踏", "Kappa/背靠背", "Lining/李宁", "nike Air Jordan/乔丹" };
        /// <summary>
        ///   将空格*/等去除，保留前六位数字
        /// </summary>
        string[] specialBrands3 = { "Adidas/阿迪达斯" };
        /// <summary>
        ///   去除字母只保留数字
        /// </summary>
        string[] specialBrands4 = { "Deerway/德尔惠" };
        /// <summary>
        ///   去除字母只保留数字，且只保留-之前的字符串
        /// </summary>
        string[] specialBrands5 = { "erke/鸿星尔克" };
        /// <summary>
        ///   首位如果有-需去除，保留-之前字符串
        /// </summary>
        string[] specialBrands6 = { "Nike/耐克" };
        /// <summary>
        ///  保留前六位数字
        /// </summary>
        string[] specialBrands7 = { "Puma/彪马" };
        /// <summary>
        ///  去除尾号字母
        /// </summary>
        string[] specialBrands8 = { "乔丹" };
        DataTable dtaxx = new DataTable();
        private string strDatetime = string.Empty;
        /// <summary>
        /// 查找sheet总个数
        /// </summary>
        private DataTable dtResultSheetCount = new DataTable();
        /// <summary>
        /// 临时使用
        /// </summary>
        private DataTable dtResultConnig = new DataTable();

        /// <summary>
        /// 构造方法
        /// </summary>
        public DataClean()
        {
            InitializeComponent();
            this.MaximizeBox = false;//禁用最大化
            this.StartPosition = FormStartPosition.CenterScreen;//窗体显示在桌面中间
            //先处理特殊品牌字符大小写
            for (int i = 0; i < specialBrands.Length; i++)
                specialBrands[i] = specialBrands[i].ToUpper().Trim();

            for (int i = 0; i < specialBrands2.Length; i++)
                specialBrands2[i] = specialBrands2[i].ToUpper().Trim();

            for (int i = 0; i < specialBrands3.Length; i++)
                specialBrands3[i] = specialBrands3[i].ToUpper().Trim();

            for (int i = 0; i < specialBrands4.Length; i++)
                specialBrands4[i] = specialBrands4[i].ToUpper().Trim();

            for (int i = 0; i < specialBrands5.Length; i++)
                specialBrands5[i] = specialBrands5[i].ToUpper().Trim();

            for (int i = 0; i < specialBrands6.Length; i++)
                specialBrands6[i] = specialBrands6[i].ToUpper().Trim();

            for (int i = 0; i < specialBrands7.Length; i++)
                specialBrands7[i] = specialBrands7[i].ToUpper().Trim();

            for (int i = 0; i < specialBrands8.Length; i++)
                specialBrands8[i] = specialBrands8[i].ToUpper().Trim();

            //Match m = Regex.Match(specialBrands[0], "[a-zA-Z]+");

            //if (m.Success)
            //{
            //    string[] acd = Regex.Split(specialBrands[1], "/");
            //}
        }

        /// <summary>
        /// 转半角的函数(DBC case)
        /// 【转换英文特殊字符串】
        ///全角空格为12288，半角空格为32
        ///其他字符半角(33-126)与全角(65281-65374)的对应关系是：均相差65248 //
        /// </summary>
        /// <param name="str">任意字符串</param>
        /// <returns>返回新字符串</returns>
        private string ToDBC(string str)
        {
            char[] oldStr = str.ToCharArray();
            for (int i = 0; i < oldStr.Length; i++)
            {
                if (oldStr[i] == 12288)
                {
                    oldStr[i] = (char)32;
                    continue;
                }
                if (oldStr[i] > 65280 && oldStr[i] < 65375)
                    oldStr[i] = (char)(oldStr[i] - 65248);
            }
            return new string(oldStr);
        }

        /// <summary>
        /// 打开文件按钮
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void butOpen_Click(object sender, EventArgs e)
        {
            string filePath = string.Empty;
            OpenFileDialog openFile = new OpenFileDialog();
            openFile.Multiselect = true;
            openFile.RestoreDirectory = true;
            path = string.Empty;
            this.button2.Text = "打开:保存excel";
            openFile.Filter = "excel07-10(*.xlsx)|*.xlsx;|excel(*.xls)|*.xls;|csv(.csv)|*.csv|(所有类型)|*.*";
            if (openFile.ShowDialog() == DialogResult.OK)
            {              
                filePath = openFile.FileName;
                string[] patht = openFile.FileNames;
                new Thread(new ThreadStart(delegate() { ListViewAddItems(patht); })).Start();
                new Thread(new ThreadStart(delegate()
                {
                    dtResultSheetCount = new DataTable();

                    dtResultSheetCount.Columns.Add("文件名");
                    dtResultSheetCount.Columns.Add("个数");

                    for (int i = 0; i < patht.Length; i++)
                    {
                        this.progressBar1.Invoke(new ThreadStart(delegate() { progressBar1.Value = 0; }));
                        this.label1.Invoke(new ThreadStart(delegate() { label2.Text = "文件个数：共 " + patht.Length + " 个文件,开始第 " + (i + 1) + " 个 "; }));
                        this.label1.Invoke(new ThreadStart(delegate() { label1.Text = "当前进度："; }));
                        //////checksum
                        ////DataTable dtxstr = new DataTable();
                        ////FileStream sr = new FileStream(patht[i], FileMode.Open,FileAccess.Read);
                      

                        ////byte[] chx = new byte[sr.Length];
                        ////sr.Read(chx, 0, chx.Length);
                        ////string txt = Encoding.GetEncoding("GBK").GetString(chx);

                        ////dtxstr.ReadXml(txt);
                        ////Debug.WriteLine(txt);
                        ////sr.Close();
                        ////StreamReader str = new StreamReader(patht[i]);
                        ////while (str.ReadLine() !=null)
                        ////{
                        ////    Debug.WriteLine(Encoding.Unicode.GetString(chx));
                        ////}
                        ////str.Close();
                        //MD5 md5 = new MD5CryptoServiceProvider();
                        //byte[] outo = md5.ComputeHash(chx);
                        //string ax = BitConverter.ToString(outo).Replace("-", "");
                        //Debug.WriteLine(ax);
                        //----------------------------------------

                        ////string asd = string.Empty;
                        ////MD5 md5 = new MD5CryptoServiceProvider();
                        ////byte[] check = new byte[1024];
                        ////MemoryStream memory = new MemoryStream();
                        ////BinaryFormatter bf = new BinaryFormatter();
                        ////bf.Serialize(memory,chx);
                        ////check = memory.GetBuffer();
                        ////memory.Close();
                        ////byte[] outo = md5.ComputeHash(check);
                        ////asd = BitConverter.ToString(outo).Replace("-", "");

                        //----------------------------------------

                        openExcel(patht[i]);
                        //查询sheet页
                        //openExcelSheets(patht[i]);
                    }
                    if (dtResultSheetCount.Rows.Count > 0)//保存
                    {
                        path = AppDomain.CurrentDomain.BaseDirectory;//路径
                        mySql.DataExportToFile(dtResultSheetCount, path + "result" + DateTime.Now.ToString().Replace("-", "").Replace(":", "").Replace("/", "") + ".xls");
                        //mySql.DataOfGetExcel(dtResultSheetCount, path + "--result" + DateTime.Now.ToString().Replace("-", "").Replace(":", "").Replace("/", "") + ".xls");
                        //测试
                        //mySql.DataExportToTextFile(dtResultSheetCount, path + "result" + DateTime.Now.Date.ToString("yyyyMMdd") + ".xls");
                        this.button2.Invoke(new ThreadStart(delegate() { this.button2.Text = path; }));
                    }
                })).Start();
            }
        }

        //加载任务列表
        /// <summary>
        /// 加载任务列表
        /// </summary>
        /// <param name="fileNames"></param>
        private void ListViewAddItems(string[] fileNames)
        {
            //开始更新数据项锁定listview控件
            listFileMessage.Invoke(new ThreadStart(delegate() { listFileMessage.BeginUpdate(); }));
            listFileMessage.Invoke(new ThreadStart(delegate() { listFileMessage.Items.Clear(); }));

            this.Invoke(new ThreadStart(delegate()
            {
                foreach (string fileName in fileNames)
                {
                    FileInfo file = new FileInfo(fileName);
                    if (listFileMessage.Items.ContainsKey(fileName))
                        continue;

                    ListViewItem lvi = new ListViewItem();
                    lvi.Name = fileName;
                    lvi.Text = fileName;
                    try
                    {
                        ListViewItem.ListViewSubItem lvsi = new ListViewItem.ListViewSubItem();
                        lvsi.Text = Convert.ToString(Math.Round((file.Length / 1024.0))) + " KB";
                        lvi.SubItems.Add(lvsi);
                        lvsi = new ListViewItem.ListViewSubItem();
                        lvsi.Text = file.LastWriteTime.ToString();
                        lvi.SubItems.Add(lvsi);
                        lvsi = new ListViewItem.ListViewSubItem();
                        lvsi.Text = string.Empty;
                        lvi.SubItems.Add(lvsi);
                        listFileMessage.Invoke(new ThreadStart(delegate() { listFileMessage.Items.Add(lvi); }));
                    }
                    catch (FileNotFoundException FFE)
                    {
                        MessageBox.Show(FFE.Message, "系统提示", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    catch (IOException IOE)
                    {
                        MessageBox.Show(IOE.Message, "系统提示", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }));
            listFileMessage.Invoke(new ThreadStart(delegate()
            {
                listFileMessage.EndUpdate();
                foreach (ColumnHeader ch in listFileMessage.Columns)
                    ch.AutoResize(ColumnHeaderAutoResizeStyle.ColumnContent);
            }));
        }

        //查询excelsheet页个数
        /// <summary>
        /// 查询excelsheet页个数
        /// </summary>
        /// <param name="filePath"></param>
        private void openExcelSheets(string filePath)
        {
            //读取excel
            //先遍历sheet个数
            List<string> workSheet = new List<string>();
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            excelApp.DisplayAlerts = false;
            excelApp.Visible = false;
            Microsoft.Office.Interop.Excel.Workbook book = excelApp.Workbooks.Open(filePath);
            foreach (Microsoft.Office.Interop.Excel.Worksheet sheet in book.Worksheets)
                workSheet.Add(sheet.Name);

            Debug.WriteLine(filePath + "--" + workSheet.Count);
            string[] fileName = filePath.Split('\\');

            DataRow dr = dtResultSheetCount.NewRow();
            dr["文件名"] = fileName[fileName.Length - 1];
            dr["个数"] = workSheet.Count;
            dtResultSheetCount.Rows.Add(dr);

            //关闭excel进程           
            try
            {
                book.Close();
                excelApp.Quit();
                if (excelApp != null)
                {
                    //获取Excel App的句柄
                    hwnd = new IntPtr(excelApp.Hwnd);
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
        /// 打开excel获取数据
        /// </summary>
        /// <param name="filePath">打开excel文件路径</param>
        private void openExcel(string filePath)
        {
            //读取excel
            //先遍历sheet个数
            List<string> workSheet = new List<string>();
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            excelApp.DisplayAlerts = false;
            excelApp.Visible = false;
            try
            {
                Microsoft.Office.Interop.Excel.Workbook book = excelApp.Workbooks.Open(filePath);
                foreach (Microsoft.Office.Interop.Excel.Worksheet sheet in book.Worksheets)
                    workSheet.Add(sheet.Name);

                strDatetime = DateTime.Now.ToString();

                DataSet Buff = new DataSet();

                using (OleDbConnection oleDbCon = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" + filePath + "';Extended Properties='Excel 12.0;HDR=Yes;IMEX=1'"))
                {
                    try
                    {
                        oleDbCon.Open();
                    }
                    catch (Exception ex)
                    {
                        throw ex;
                    }
                    //遍历所有表
                    foreach (string sheetName in workSheet)
                    {
                        Debug.WriteLine(string.Format("正在读取数据源工作表--{0}--处理函数：{1}", sheetName, DateTime.Now.ToString()));
                        //读取EXCEL数据源
                        //string cmd = "SELECT 商品id,商品名称,店铺名称,品牌名称,品牌id,品类id,品类名称,吊牌价,运动鞋分类,款号,鞋帮高度,闭合方式,性别  FROM [" + sheetName + "$] order by 品牌名称,款号,吊牌价";  
                        ////女装
                        ////select 平台,商品名称,urlID,skuId,品牌ID,品牌,品类id,品类1,品类2,店铺,店铺ID,标牌价,促销信息,累积评论数,商品评分,月成交记录
                        // string cmd = "SELECT 平台,商品名称,urlID,skuId,品牌ID,品牌,品类id,品类1,品类2,店铺,店铺ID,标牌价,促销信息,累积评论数,商品评分,月成交记录,衣长,袖长,衣门襟,图案,风格,货号,里料材质,厚薄,领子,促销价 FROM [" + sheetName + "$] ";
                        string cmd = "SELECT * FROM [" + sheetName + "$] ";
                        OleDbDataAdapter adp = new OleDbDataAdapter(cmd, oleDbCon);
                        adp.SelectCommand.CommandTimeout = 3600;
                        try
                        {
                            DataTable data = new DataTable();
                            adp.Fill(data);
                            if (data != null && data.Rows.Count > 0)
                            {
                                //string asd = string.Empty;
                                //MD5 md5 = new MD5CryptoServiceProvider();
                                //byte[] check = new byte[1024];
                                //MemoryStream memory = new MemoryStream();
                                //BinaryFormatter bf = new BinaryFormatter();
                                //bf.Serialize(memory, data);
                                //check = memory.GetBuffer();
                                //memory.Close();
                                //byte[] outo = md5.ComputeHash(check);
                                //asd = BitConverter.ToString(outo).Replace("-", "");

                                //Debug.WriteLine(asd);



                                Debug.WriteLine("star:" + DateTime.Now.ToString());

                                data = data.AsEnumerable().Select(
                                    p =>
                                    {
                                        foreach (DataColumn dc in data.Columns)
                                        {
                                            if (p[dc].ToString() == "")
                                                p[dc] = DBNull.Value;
                                        }
                                        return p;
                                    }
                                    ).CopyToDataTable();
                                Debug.WriteLine("end:" + DateTime.Now.ToString());
                                // openExcelSheets(data, filePath);
                                //SelectYJBLowerPrice(data, filePath);
                                //////执行第一版方法
                                ////TheFirstMethod(data, sheetName);
                                ////执行第二版方法
                                //// TheSecondMethod(data, sheetName);
                                //// TheThirdMethod(data, sheetName);
                                ////TheFRMethod(data, sheetName);
                                int dpBrand = 0;
                                for (int ax = 0; ax < data.Columns.Count; ax++)
                                {
                                    if (data.Columns[ax].ColumnName == "单品品牌")
                                    {
                                        dpBrand = 1;
                                        break;
                                    }
                                }
                                if (dpBrand == 1)
                                {
                                    Online_TC(data, sheetName);
                                }
                                else
                                {
                                    //整理套餐机型
                                    TC(data, sheetName);
                                }
                                ////女装
                                ////NewMethodNV(data, sheetName);
                                //// TddC(data, sheetName);
                            }
                        }
                        catch (Exception ex)
                        {
                            throw ex;
                        }
                    }
                }
                //关闭excel进程           
                try
                {
                    book.Close();
                    excelApp.Quit();
                    if (excelApp != null)
                    {
                        //获取Excel App的句柄
                        hwnd = new IntPtr(excelApp.Hwnd);
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
            catch (Exception e)
            {
                MessageBox.Show("error:" + e.Message ); return;
            }
        }


        private void ProgressBarDisplay(DataTable data, int i)
        {
            progressBar1.Value = (i + 1) * 100 / (data.Rows.Count);
            label1.Text = "当前进度：" + (i + 1) + "/" + (data.Rows.Count) + "";
            if (i > 1000)
            {
                if (i % 1000 == 0)
                    GC.Collect();
            }
            Debug.WriteLine((i + 1) + "：" + DateTime.Now.ToString());
        }

        private void butDownload_Click(object sender, EventArgs e)
        {
            FtpClass ftpapp = new FtpClass(userModel.FtpIp, userModel.FtpUserName, userModel.FtpUserPwd);//打开ftp
            //下载数据
            ftpapp.Download(@"\AVC CleanData\AVC CleanData\bin", "/Report/常规报告/冰柜/线上/月报/14.12/AVC-冰柜-线上零售市场监测月度数据报告（14.12）.xlsx", true, ref message);
        }

        private void butOpentxtFile_Click(object sender, EventArgs e)
        {
            string filePath = string.Empty;
            OpenFileDialog openFile = new OpenFileDialog();
            openFile.Multiselect = true;
            openFile.RestoreDirectory = true;
            openFile.Filter = "文本文件(*.txt)|*.txt;";
            //openFile.Filter = "excel(*.xls)|*.xls;|excel07-10(*.xlsx)|*.xlsx";
            //openFile.InitialDirectory = @"F:\安装文件\新建文件夹";
            if (openFile.ShowDialog() == DialogResult.OK)
            {
                filePath = openFile.FileName;
                string[] path = openFile.FileNames;
                new Thread(new ThreadStart(delegate() { ListViewAddItems(path); })).Start();
                new Thread(new ThreadStart(delegate()
                {
                    for (int i = 0; i < path.Length; i++)
                    {
                        this.progressBar1.Invoke(new ThreadStart(delegate() { progressBar1.Value = 0; }));
                        this.label1.Invoke(new ThreadStart(delegate() { label2.Text = "文件个数：共 " + path.Length + " 个文件,开始第 " + (i + 1) + " 个 "; }));
                        this.label1.Invoke(new ThreadStart(delegate() { label1.Text = "当前进度："; }));
                        openTxtFile(path[i]);
                    }
                })).Start();
            }
        }

        private void openTxtFile(string path)
        {
            StreamWriter strw = new StreamWriter(@"C:\Users\Administrator\Desktop\wgh.txt", true, Encoding.UTF8);

            StreamReader stread = new StreamReader(path, Encoding.Default);

            while ((stread.ReadLine()) != null)
            {
                strw.WriteLine(stread.ReadLine());
                //Debug.WriteLine(stread.ReadLine());
            }
            stread.Close();
            strw.Close();
            MessageBox.Show("结束", "提示");
        }

        private void butSelectData_Click(object sender, EventArgs e)
        {
            SaveFileDialog saveFile = new SaveFileDialog();
            saveFile.Title = "保存为Excel文件";
            saveFile.Filter = "07-10Excel工作薄 (*.xlsx)|*.xlsx";
            saveFile.RestoreDirectory = true;
            if (saveFile.ShowDialog() == DialogResult.OK)
            {
                //mySql.DataExportToFile(exportDs.Tables["exportDs"], saveFile.FileName,label1, progressBar1);
            }
        }

        /// <summary>
        /// 第一版 计算 方法
        /// </summary>
        /// <param name="dtTable"></param>
        private void TheFirstMethod(DataTable dtTable, string tableName)
        {
            //商品id、商品名称、店铺名称、品牌id、品牌名称、品类id、品类名称、吊牌价、运动鞋分类 
            #region 测试（复杂方法）
            DataTable dtBiaozhunBrand = new DataTable();//存放标准品牌
            DataTable dtBiaozhunData = new DataTable();//存放查找到的标准数据
            //加载列名
            for (int i = 0; i < dtTable.Columns.Count; i++)
            {
                dtBiaozhunBrand.Columns.Add(dtTable.Columns[i].ColumnName);
                dtBiaozhunData.Columns.Add(dtTable.Columns[i].ColumnName);
            }
            //dtBiaozhunBrand = dtTable.Copy();
            //dtBiaozhunBrand、dtBiaozhunData在添加一个 标准字段，avc自定义（avc商品id、avc品牌id）
            dtBiaozhunBrand.Columns.Add("avc商品id", typeof(string));
            dtBiaozhunBrand.Columns.Add("avc品牌id", typeof(string));
            dtBiaozhunData.Columns.Add("avc商品id", typeof(string));
            dtBiaozhunData.Columns.Add("avc品牌id", typeof(string));

            for (int i = 0; i < dtTable.Rows.Count; i++)
            {
                if (i == 0)//第一条数据默认为标准数据
                {
                    dtBiaozhunBrand.Rows.Add(dtTable.Rows[i].ItemArray);//直接复制数据 行 
                    dtBiaozhunBrand.Rows[dtBiaozhunBrand.Rows.Count - 1]["avc商品id"] = "avc00_" + i + "";
                    dtBiaozhunBrand.Rows[dtBiaozhunBrand.Rows.Count - 1]["avc品牌id"] = "avcb00_" + i + "";
                    dtBiaozhunData.Rows.Add(dtTable.Rows[i].ItemArray);
                    dtBiaozhunData.Rows[dtBiaozhunData.Rows.Count - 1]["avc商品id"] = "avc00_" + i + "";
                    dtBiaozhunData.Rows[dtBiaozhunData.Rows.Count - 1]["avc品牌id"] = "avcb00_" + i + "";
                    continue;//继续下条数据
                }
                //开始逐条数据查找判断
                #region （从第二条数据开始） 开始逐条数据查找判断
                int need = 1;//查看该条数据是否在标准列表里面查找到.
                for (int a = 0; a < dtBiaozhunBrand.Rows.Count; a++)
                {
                    if (dtTable.Rows[i]["品牌id"].ToString().ToUpper().Trim() == dtBiaozhunBrand.Rows[a]["品牌id"].ToString().ToUpper().Trim())//第一步（品牌相等 ）进行下一步
                    {
                        if (dtTable.Rows[i]["款号"].ToString().ToUpper().Trim() == dtBiaozhunBrand.Rows[a]["款号"].ToString().ToUpper().Trim())//第二步（款号相等 ）进行下一步
                        {
                            //开始计算 属性匹配度                                            
                            //属性相同=1，不同=0，NULL=空
                            //得分=0.25*吊牌价匹配+0.25*鞋帮高度匹配+0.25*闭合方式+0.25*性别
                            decimal codePrice = 0, codeHigh = 0, codeClose = 0, codeSex = 0;
                            if (dtTable.Rows[i]["鞋帮高度"].ToString().ToUpper().Trim() == dtBiaozhunBrand.Rows[a]["鞋帮高度"].ToString().ToUpper().Trim())
                                codeHigh = 1;
                            if (dtTable.Rows[i]["闭合方式"].ToString().ToUpper().Trim() == dtBiaozhunBrand.Rows[a]["闭合方式"].ToString().ToUpper().Trim())
                                codeClose = 1;
                            if (dtTable.Rows[i]["性别"].ToString().ToUpper().Trim() == dtBiaozhunBrand.Rows[a]["性别"].ToString().ToUpper().Trim())
                                codeSex = 1;
                            if (dtTable.Rows[i]["吊牌价"].ToString().ToUpper().Trim() == dtBiaozhunBrand.Rows[a]["吊牌价"].ToString().ToUpper().Trim())
                                codePrice = 1;
                            if (dtTable.Rows[i]["吊牌价"].ToString().ToUpper().Trim() != "NULL" && dtTable.Rows[i]["吊牌价"].ToString().ToUpper().Trim() != "")
                            {
                                if (dtBiaozhunBrand.Rows[a]["吊牌价"].ToString().ToUpper().Trim() != "NULL" && dtBiaozhunBrand.Rows[a]["吊牌价"].ToString().ToUpper().Trim() != "")
                                {
                                    if (Convert.ToDecimal(dtTable.Rows[i]["吊牌价"].ToString().ToUpper().Trim()) > Convert.ToDecimal(dtBiaozhunBrand.Rows[a]["吊牌价"].ToString().ToUpper().Trim()))
                                    {
                                        if (Convert.ToDecimal(dtBiaozhunBrand.Rows[a]["吊牌价"].ToString().ToUpper().Trim()) * 100 / Convert.ToDecimal(dtTable.Rows[i]["吊牌价"].ToString().ToUpper().Trim()) >= 90)
                                            codePrice = 1;
                                    }
                                    else if (Convert.ToDecimal(dtTable.Rows[i]["吊牌价"].ToString().ToUpper().Trim()) < Convert.ToDecimal(dtBiaozhunBrand.Rows[a]["吊牌价"].ToString().ToUpper().Trim()))
                                    {
                                        if (Convert.ToDecimal(dtTable.Rows[i]["吊牌价"].ToString().ToUpper().Trim()) * 100 / Convert.ToDecimal(dtBiaozhunBrand.Rows[a]["吊牌价"].ToString().ToUpper().Trim()) >= 90)
                                            codePrice = 1;
                                    }
                                }
                            }
                            //最终匹配度
                            decimal endCode = Convert.ToDecimal(0.25) * (codePrice + codeHigh + codeClose + codeSex);
                            if (endCode >= Convert.ToDecimal(0.95))
                            {
                                need = 0;
                                dtBiaozhunData.Rows.Add(dtTable.Rows[i].ItemArray);
                                dtBiaozhunData.Rows[dtBiaozhunData.Rows.Count - 1]["avc商品id"] = dtBiaozhunBrand.Rows[a]["avc商品id"];
                                dtBiaozhunData.Rows[dtBiaozhunData.Rows.Count - 1]["avc品牌id"] = dtBiaozhunBrand.Rows[a]["avc品牌id"];
                                break;//跳出查找循环
                            }
                            else if (endCode >= Convert.ToDecimal(0.75))//匹配度>=0.75 则 判断（品类名称）是否相等
                            {
                                if (dtTable.Rows[i]["品类名称"].ToString().ToUpper().Trim() == dtBiaozhunBrand.Rows[a]["品类名称"].ToString().ToUpper().Trim())//第三步（品类名称相等）
                                {
                                    need = 0;
                                    dtBiaozhunData.Rows.Add(dtTable.Rows[i].ItemArray);
                                    dtBiaozhunData.Rows[dtBiaozhunData.Rows.Count - 1]["avc商品id"] = dtBiaozhunBrand.Rows[a]["avc商品id"];
                                    dtBiaozhunData.Rows[dtBiaozhunData.Rows.Count - 1]["avc品牌id"] = dtBiaozhunBrand.Rows[a]["avc品牌id"];
                                    break;//跳出查找循环
                                }
                            }
                        }
                    }
                }
                if (need == 1)
                {
                    dtBiaozhunBrand.Rows.Add(dtTable.Rows[i].ItemArray);//添加为新标准品牌
                    dtBiaozhunBrand.Rows[dtBiaozhunBrand.Rows.Count - 1]["avc商品id"] = "avc00_" + i + "";
                    dtBiaozhunBrand.Rows[dtBiaozhunBrand.Rows.Count - 1]["avc品牌id"] = "avcb00_" + i + "";
                    //添加为标准数据
                    dtBiaozhunData.Rows.Add(dtTable.Rows[i].ItemArray);
                    dtBiaozhunData.Rows[dtBiaozhunData.Rows.Count - 1]["avc商品id"] = "avc00_" + i + "";
                    dtBiaozhunData.Rows[dtBiaozhunData.Rows.Count - 1]["avc品牌id"] = "avcb00_" + i + "";
                }
                #endregion
                this.Invoke(new ThreadStart(delegate()
                {
                    ProgressBarDisplay(dtTable, i);
                }));
            }
            //导出数据
            // mySql.DataExportToFile(dtBiaozhunData, @"wgh" + tableName + ".xls");

            #endregion
        }

        /// <summary>
        /// 第二版-------------
        /// </summary>
        /// <param name="dtTable"></param>
        /// <param name="tableName"></param>
        private void TheSecondMethod(DataTable dtTable, string tableName)
        {
            //商品id、商品名称、店铺名称、品牌id、品牌名称、品类id、品类名称、吊牌价、运动鞋分类 
            #region 测试（复杂方法）
            DataTable dtBiaozhunBrand = new DataTable();//存放标准品牌
            DataTable dtBiaozhunData = new DataTable();//存放查找到的标准数据
            //加载列名
            for (int i = 0; i < dtTable.Columns.Count; i++)
            {
                dtBiaozhunBrand.Columns.Add(dtTable.Columns[i].ColumnName);
                dtBiaozhunData.Columns.Add(dtTable.Columns[i].ColumnName);
            }
            //dtBiaozhunBrand = dtTable.Copy();
            //dtBiaozhunBrand、dtBiaozhunData在添加一个 标准字段，avc自定义（avc商品id、avc品牌id）
            dtBiaozhunBrand.Columns.Add("avc商品id", typeof(string));
            dtBiaozhunBrand.Columns.Add("avc品牌id", typeof(string));
            dtBiaozhunData.Columns.Add("avc商品id", typeof(string));
            dtBiaozhunData.Columns.Add("avc品牌id", typeof(string));

            for (int i = 0; i < dtTable.Rows.Count; i++)
            {
                if (i == 0 && !specialBrands.Contains(dtTable.Rows[i]["品牌名称"].ToString().ToUpper().Trim()))//第一条数据默认为标准数据
                {
                    dtBiaozhunBrand.Rows.Add(dtTable.Rows[i].ItemArray);//直接复制数据 行 
                    dtBiaozhunBrand.Rows[dtBiaozhunBrand.Rows.Count - 1]["avc商品id"] = "avc00_" + i + "";
                    dtBiaozhunBrand.Rows[dtBiaozhunBrand.Rows.Count - 1]["avc品牌id"] = "avcb00_" + i + "";
                    dtBiaozhunData.Rows.Add(dtTable.Rows[i].ItemArray);
                    dtBiaozhunData.Rows[dtBiaozhunData.Rows.Count - 1]["avc商品id"] = "avc00_" + i + "";
                    dtBiaozhunData.Rows[dtBiaozhunData.Rows.Count - 1]["avc品牌id"] = "avcb00_" + i + "";
                    continue;//继续下条数据
                }

                //开始逐条数据查找判断
                #region （从第二条数据开始） 开始逐条数据查找判断
                int need = 1;//查看该条数据是否在标准列表里面查找到.
                for (int a = 0; a < dtBiaozhunBrand.Rows.Count; a++)
                {
                    #region //非特殊品牌（使用属性不同标记分割，吊牌价区别）

                    if (!specialBrands.Contains(dtTable.Rows[i]["品牌名称"].ToString().ToUpper().Trim()))
                    {

                        if (dtTable.Rows[i]["品牌id"].ToString().ToUpper().Trim() == dtBiaozhunBrand.Rows[a]["品牌id"].ToString().ToUpper().Trim())
                        {
                            if (dtTable.Rows[i]["吊牌价"].ToString().ToUpper().Trim() != "NULL" && dtTable.Rows[i]["吊牌价"].ToString().ToUpper().Trim() != "")
                            {
                                if (dtBiaozhunBrand.Rows[a]["吊牌价"].ToString().ToUpper().Trim() != "NULL" && dtBiaozhunBrand.Rows[a]["吊牌价"].ToString().ToUpper().Trim() != "")
                                {
                                    if (Convert.ToDecimal(dtTable.Rows[i]["吊牌价"].ToString().ToUpper().Trim()) == Convert.ToDecimal(dtBiaozhunBrand.Rows[a]["吊牌价"].ToString().ToUpper().Trim()))
                                    {
                                        if (Approximate(dtTable.Rows[i]["款号"].ToString().ToUpper().Trim(), dtBiaozhunBrand.Rows[a]["款号"].ToString().Trim()) >= 91)
                                        {
                                            need = 0;
                                            dtBiaozhunData.Rows.Add(dtTable.Rows[i].ItemArray);
                                            dtBiaozhunData.Rows[dtBiaozhunData.Rows.Count - 1]["avc商品id"] = dtBiaozhunBrand.Rows[a]["avc商品id"];
                                            dtBiaozhunData.Rows[dtBiaozhunData.Rows.Count - 1]["avc品牌id"] = dtBiaozhunBrand.Rows[a]["avc品牌id"];
                                            break;//跳出查找循环
                                        }
                                    }
                                }
                            }
                        }
                    }

                    #endregion

                    #region //特殊品牌处理

                    #region 款号只保留-之前的字符串
                    else
                    {
                        string[] GirardData = null;
                        string[] GirardBrand = null;
                        if (specialBrands2.Contains(dtTable.Rows[i]["品牌名称"].ToString().ToUpper().Trim()))//款号只保留-之前的字符串
                        {
                            //款号截取
                            GirardData = Regex.Split(dtTable.Rows[i]["款号"].ToString().ToUpper().Trim(), "-");
                            //标准款号截取
                            GirardBrand = Regex.Split(dtBiaozhunBrand.Rows[a]["款号"].ToString().ToUpper().Trim(), "-");
                        }
                        else if (specialBrands3.Contains(dtTable.Rows[i]["品牌名称"].ToString().ToUpper().Trim()))//将空格*/等去除，保留前六位数字（Adidas/阿迪达斯：2015SSOR-KCO11 款式格式为此，不参与）
                        {
                            if (dtTable.Rows[i]["款号"].ToString().ToUpper().Trim().Contains("-") && dtTable.Rows[i]["款号"].ToString().Trim().Contains("201"))
                            {
                                if (dtTable.Rows[i]["款号"].ToString().ToUpper().Trim() == dtBiaozhunBrand.Rows[a]["款号"].ToString().ToUpper().Trim())//第一步 ：款号 相等
                                {
                                    if (dtTable.Rows[i]["品类id"].ToString().ToUpper().Trim() == dtBiaozhunBrand.Rows[a]["品类id"].ToString().ToUpper().Trim())//第二步：品类名称 相等 属性匹配度>=0.75
                                    {
                                        //开始计算 属性匹配度                                            
                                        //属性相同=1，不同=0，NULL=空
                                        //得分=0.25*吊牌价匹配+0.25*鞋帮高度匹配+0.25*闭合方式+0.25*性别
                                        decimal codePrice = 0, codeHigh = 0, codeClose = 0, codeSex = 0;
                                        if (dtTable.Rows[i]["鞋帮高度"].ToString().ToUpper().Trim() == dtBiaozhunBrand.Rows[a]["鞋帮高度"].ToString().ToUpper().Trim())
                                            codeHigh = 1;
                                        if (dtTable.Rows[i]["闭合方式"].ToString().ToUpper().Trim() == dtBiaozhunBrand.Rows[a]["闭合方式"].ToString().ToUpper().Trim())
                                            codeClose = 1;
                                        if (dtTable.Rows[i]["性别"].ToString().ToUpper().Trim() == dtBiaozhunBrand.Rows[a]["性别"].ToString().ToUpper().Trim())
                                            codeSex = 1;
                                        if (dtTable.Rows[i]["吊牌价"].ToString().ToUpper().Trim() == dtBiaozhunBrand.Rows[a]["吊牌价"].ToString().ToUpper().Trim())
                                            codePrice = 1;
                                        if (dtTable.Rows[i]["吊牌价"].ToString().ToUpper().Trim() != "NULL" && dtTable.Rows[i]["吊牌价"].ToString().ToUpper().Trim() != "")
                                        {
                                            if (dtBiaozhunBrand.Rows[a]["吊牌价"].ToString().ToUpper().Trim() != "NULL" && dtBiaozhunBrand.Rows[a]["吊牌价"].ToString().ToUpper().Trim() != "")
                                            {
                                                if (Convert.ToDecimal(dtTable.Rows[i]["吊牌价"].ToString().ToUpper().Trim()) > Convert.ToDecimal(dtBiaozhunBrand.Rows[a]["吊牌价"].ToString().ToUpper().Trim()))
                                                {
                                                    if (Convert.ToDecimal(dtBiaozhunBrand.Rows[a]["吊牌价"].ToString().ToUpper().Trim()) * 100 / Convert.ToDecimal(dtTable.Rows[i]["吊牌价"].ToString().ToUpper().Trim()) >= 90)
                                                        codePrice = 1;
                                                }
                                                else if (Convert.ToDecimal(dtTable.Rows[i]["吊牌价"].ToString().ToUpper().Trim()) < Convert.ToDecimal(dtBiaozhunBrand.Rows[a]["吊牌价"].ToString().ToUpper().Trim()))
                                                {
                                                    if (Convert.ToDecimal(dtTable.Rows[i]["吊牌价"].ToString().ToUpper().Trim()) * 100 / Convert.ToDecimal(dtBiaozhunBrand.Rows[a]["吊牌价"].ToString().ToUpper().Trim()) >= 90)
                                                        codePrice = 1;
                                                }
                                            }
                                        }
                                        //最终匹配度
                                        decimal endCode = Convert.ToDecimal(0.25) * (codePrice + codeHigh + codeClose + codeSex);
                                        if (endCode >= Convert.ToDecimal(0.75))//匹配度>=0.75 则 判断（品类名称）是否相等
                                        {
                                            if (dtTable.Rows[i]["品类名称"].ToString().ToUpper().Trim() == dtBiaozhunBrand.Rows[a]["品类名称"].ToString().ToUpper().Trim())//第三步（品类名称相等）
                                            {
                                                need = 0;
                                                dtBiaozhunData.Rows.Add(dtTable.Rows[i].ItemArray);
                                                dtBiaozhunData.Rows[dtBiaozhunData.Rows.Count - 1]["avc商品id"] = dtBiaozhunBrand.Rows[a]["avc商品id"];
                                                dtBiaozhunData.Rows[dtBiaozhunData.Rows.Count - 1]["avc品牌id"] = dtBiaozhunBrand.Rows[a]["avc品牌id"];
                                                break;//跳出查找循环
                                            }
                                        }
                                    }
                                    else//品类名称不相等 属性匹配度=1
                                    {
                                        //开始计算 属性匹配度                                            
                                        //属性相同=1，不同=0，NULL=空
                                        //得分=0.25*吊牌价匹配+0.25*鞋帮高度匹配+0.25*闭合方式+0.25*性别
                                        decimal codePrice = 0, codeHigh = 0, codeClose = 0, codeSex = 0;
                                        if (dtTable.Rows[i]["鞋帮高度"].ToString().ToUpper().Trim() == dtBiaozhunBrand.Rows[a]["鞋帮高度"].ToString().ToUpper().Trim())
                                            codeHigh = 1;
                                        if (dtTable.Rows[i]["闭合方式"].ToString().ToUpper().Trim() == dtBiaozhunBrand.Rows[a]["闭合方式"].ToString().ToUpper().Trim())
                                            codeClose = 1;
                                        if (dtTable.Rows[i]["性别"].ToString().ToUpper().Trim() == dtBiaozhunBrand.Rows[a]["性别"].ToString().ToUpper().Trim())
                                            codeSex = 1;
                                        if (dtTable.Rows[i]["吊牌价"].ToString().ToUpper().Trim() == dtBiaozhunBrand.Rows[a]["吊牌价"].ToString().ToUpper().Trim())
                                            codePrice = 1;
                                        if (dtTable.Rows[i]["吊牌价"].ToString().ToUpper().Trim() != "NULL" && dtTable.Rows[i]["吊牌价"].ToString().ToUpper().Trim() != "")
                                        {
                                            if (dtBiaozhunBrand.Rows[a]["吊牌价"].ToString().ToUpper().Trim() != "NULL" && dtBiaozhunBrand.Rows[a]["吊牌价"].ToString().ToUpper().Trim() != "")
                                            {
                                                if (Convert.ToDecimal(dtTable.Rows[i]["吊牌价"].ToString().ToUpper().Trim()) > Convert.ToDecimal(dtBiaozhunBrand.Rows[a]["吊牌价"].ToString().ToUpper().Trim()))
                                                {
                                                    if (Convert.ToDecimal(dtBiaozhunBrand.Rows[a]["吊牌价"].ToString().ToUpper().Trim()) * 100 / Convert.ToDecimal(dtTable.Rows[i]["吊牌价"].ToString().ToUpper().Trim()) >= 90)
                                                        codePrice = 1;
                                                }
                                                else if (Convert.ToDecimal(dtTable.Rows[i]["吊牌价"].ToString().ToUpper().Trim()) < Convert.ToDecimal(dtBiaozhunBrand.Rows[a]["吊牌价"].ToString().ToUpper().Trim()))
                                                {
                                                    if (Convert.ToDecimal(dtTable.Rows[i]["吊牌价"].ToString().ToUpper().Trim()) * 100 / Convert.ToDecimal(dtBiaozhunBrand.Rows[a]["吊牌价"].ToString().ToUpper().Trim()) >= 90)
                                                        codePrice = 1;
                                                }
                                            }
                                        }
                                        //最终匹配度
                                        decimal endCode = Convert.ToDecimal(0.25) * (codePrice + codeHigh + codeClose + codeSex);
                                        if (endCode == Convert.ToDecimal(1))//匹配度==1 则 判断（品类名称）是否相等
                                        {
                                            if (dtTable.Rows[i]["品类名称"].ToString().ToUpper().Trim() == dtBiaozhunBrand.Rows[a]["品类名称"].ToString().ToUpper().Trim())//第三步（品类名称相等）
                                            {
                                                need = 0;
                                                dtBiaozhunData.Rows.Add(dtTable.Rows[i].ItemArray);
                                                dtBiaozhunData.Rows[dtBiaozhunData.Rows.Count - 1]["avc商品id"] = dtBiaozhunBrand.Rows[a]["avc商品id"];
                                                dtBiaozhunData.Rows[dtBiaozhunData.Rows.Count - 1]["avc品牌id"] = dtBiaozhunBrand.Rows[a]["avc品牌id"];
                                                break;//跳出查找循环
                                            }
                                        }
                                    }
                                }
                                continue;
                            }
                            string girardData1 = string.Empty;
                            string girardData2 = string.Empty;
                            girardData1 = Regex.Replace(dtTable.Rows[i]["款号"].ToString().ToUpper().Trim(), @"[^a-zA-Z0-9\u4E00-\u9FA5\uF900-\uFA2D]", "").Replace(" ", "");
                            girardData2 = Regex.Replace(dtBiaozhunBrand.Rows[a]["款号"].ToString().ToUpper().Trim(), @"[^a-zA-Z0-9\u4E00-\u9FA5\uF900-\uFA2D]", "").Replace(" ", "");
                            if (girardData1.Length > 6)
                                girardData1 = girardData1.Substring(0, 6);
                            if (girardData2.Length > 6)
                                girardData2 = girardData2.Substring(0, 6);
                            //款号截取
                            GirardData = Regex.Split(girardData1, "~");
                            //标准款号截取
                            GirardBrand = Regex.Split(girardData2, "~");
                        }
                        else if (specialBrands4.Contains(dtTable.Rows[i]["品牌名称"].ToString().ToUpper().Trim()))//去除字母只保留数字
                        {
                            string girardData1 = string.Empty;
                            string girardData2 = string.Empty;
                            Regex regex = new Regex(@"[0-9]+");//所有数字
                            girardData1 = regex.Match(dtTable.Rows[i]["款号"].ToString().ToUpper().Trim()).ToString();
                            girardData2 = regex.Match(dtBiaozhunBrand.Rows[a]["款号"].ToString().ToUpper().Trim()).ToString();
                            //款号截取
                            GirardData = Regex.Split(girardData1, "~");
                            //标准款号截取
                            GirardBrand = Regex.Split(girardData2, "~");
                        }
                        else if (specialBrands5.Contains(dtTable.Rows[i]["品牌名称"].ToString().ToUpper().Trim()))//去除字母只保留数字，且只保留-之前的字符串
                        {
                            string girardData1 = string.Empty;
                            string girardData2 = string.Empty;
                            Regex regex = new Regex(@"[0-9]+");//所有数字
                            girardData1 = regex.Match(dtTable.Rows[i]["款号"].ToString().ToUpper().Trim()).ToString();
                            girardData2 = regex.Match(dtBiaozhunBrand.Rows[a]["款号"].ToString().ToUpper().Trim()).ToString();

                            GirardData = Regex.Split(girardData1, "-");

                            GirardBrand = Regex.Split(girardData2, "-");
                        }
                        else if (specialBrands6.Contains(dtTable.Rows[i]["品牌名称"].ToString().ToUpper().Trim()))//首位如果有-需去除，保留-之前字符串
                        {
                            string girardData1 = string.Empty;
                            string girardData2 = string.Empty;
                            if (dtTable.Rows[i]["款号"].ToString().ToUpper().Trim()[0] == '-')
                                girardData1 = dtTable.Rows[i]["款号"].ToString().ToUpper().Trim().Substring(1);//清洗第一个 '-'
                            else
                                girardData1 = dtTable.Rows[i]["款号"].ToString().ToUpper().Trim();

                            if (dtBiaozhunBrand.Rows[a]["款号"].ToString().ToUpper().Trim()[0] == '-')
                                girardData2 = dtBiaozhunBrand.Rows[a]["款号"].ToString().ToUpper().Trim().Substring(1);//清洗第一个 '-'
                            else
                                girardData2 = dtBiaozhunBrand.Rows[a]["款号"].ToString().ToUpper().Trim();
                            //款号截取
                            GirardData = Regex.Split(girardData1, "-");
                            //标准款号截取
                            GirardBrand = Regex.Split(girardData2, "-");
                        }
                        else if (specialBrands7.Contains(dtTable.Rows[i]["品牌名称"].ToString().ToUpper().Trim()))//保留前六位数字
                        {
                            string girardData1 = string.Empty;
                            string girardData2 = string.Empty;
                            Regex regex = new Regex(@"[0-9]+");//所有数字
                            girardData1 = regex.Match(dtTable.Rows[i]["款号"].ToString().ToUpper().Trim()).ToString();
                            girardData2 = regex.Match(dtBiaozhunBrand.Rows[a]["款号"].ToString().ToUpper().Trim()).ToString();
                            //判断是否款号字符长度>6
                            if (girardData1.Length > 6)
                                girardData1 = girardData1.Substring(0, 6);
                            if (girardData2.Length > 6)
                                girardData2 = girardData2.Substring(0, 6);
                            //款号截取
                            GirardData = Regex.Split(girardData1, "~");
                            //标准款号截取
                            GirardBrand = Regex.Split(girardData2, "~");
                        }
                        else if (specialBrands8.Contains(dtTable.Rows[i]["品牌名称"].ToString().ToUpper().Trim()))//去除尾号字母
                        {
                            string girardData1 = string.Empty;
                            string girardData2 = string.Empty;
                            string haveabc = "";
                            Regex regex = new Regex(@"\d[a-zA-Z]");//以字母结尾
                            MatchCollection mm1 = regex.Matches(dtTable.Rows[i]["款号"].ToString().ToUpper().Trim());
                            if (mm1.Count > 0)
                            {
                                haveabc = mm1[mm1.Count - 1].Value.ToString();
                                girardData1 = dtTable.Rows[i]["款号"].ToString().ToUpper().Trim().Substring(0, dtTable.Rows[i]["款号"].ToString().ToUpper().Trim().LastIndexOf(haveabc) + 1);
                                haveabc = "";
                            }
                            else
                                girardData1 = dtTable.Rows[i]["款号"].ToString().ToUpper().Trim();

                            MatchCollection mm2 = regex.Matches(dtBiaozhunBrand.Rows[a]["款号"].ToString().ToUpper().Trim());
                            if (mm2.Count > 0)
                            {
                                haveabc = mm2[mm2.Count - 1].Value.ToString();
                                girardData2 = dtBiaozhunBrand.Rows[a]["款号"].ToString().ToUpper().Trim().Substring(0, dtBiaozhunBrand.Rows[a]["款号"].ToString().ToUpper().Trim().LastIndexOf(haveabc) + 1);
                            }
                            else
                                girardData2 = dtBiaozhunBrand.Rows[a]["款号"].ToString().ToUpper().Trim();
                            //款号截取
                            GirardData = Regex.Split(girardData1, "~");
                            //标准款号截取
                            GirardBrand = Regex.Split(girardData2, "~");
                        }
                        if (GirardBrand[0].ToUpper().Trim() == GirardData[0].ToUpper().Trim())
                        {
                            if (dtTable.Rows[i]["品类id"].ToString().ToUpper().Trim() == dtBiaozhunBrand.Rows[a]["品类id"].ToString().ToUpper().Trim())//第二步：品类名称 相等 属性匹配度>=0.75
                            {
                                //开始计算 属性匹配度                                            
                                //属性相同=1，不同=0，NULL=空
                                //得分=0.25*吊牌价匹配+0.25*鞋帮高度匹配+0.25*闭合方式+0.25*性别
                                decimal codePrice = 0, codeHigh = 0, codeClose = 0, codeSex = 0;
                                if (dtTable.Rows[i]["鞋帮高度"].ToString().ToUpper().Trim() == dtBiaozhunBrand.Rows[a]["鞋帮高度"].ToString().ToUpper().Trim())
                                    codeHigh = 1;
                                if (dtTable.Rows[i]["闭合方式"].ToString().ToUpper().Trim() == dtBiaozhunBrand.Rows[a]["闭合方式"].ToString().ToUpper().Trim())
                                    codeClose = 1;
                                if (dtTable.Rows[i]["性别"].ToString().ToUpper().Trim() == dtBiaozhunBrand.Rows[a]["性别"].ToString().ToUpper().Trim())
                                    codeSex = 1;
                                if (dtTable.Rows[i]["吊牌价"].ToString().ToUpper().Trim() == dtBiaozhunBrand.Rows[a]["吊牌价"].ToString().ToUpper().Trim())
                                    codePrice = 1;
                                if (dtTable.Rows[i]["吊牌价"].ToString().ToUpper().Trim() != "NULL" && dtTable.Rows[i]["吊牌价"].ToString().ToUpper().Trim() != "")
                                {
                                    if (dtBiaozhunBrand.Rows[a]["吊牌价"].ToString().ToUpper().Trim() != "NULL" && dtBiaozhunBrand.Rows[a]["吊牌价"].ToString().ToUpper().Trim() != "")
                                    {
                                        if (Convert.ToDecimal(dtTable.Rows[i]["吊牌价"].ToString().ToUpper().Trim()) > Convert.ToDecimal(dtBiaozhunBrand.Rows[a]["吊牌价"].ToString().ToUpper().Trim()))
                                        {
                                            if (Convert.ToDecimal(dtBiaozhunBrand.Rows[a]["吊牌价"].ToString().ToUpper().Trim()) * 100 / Convert.ToDecimal(dtTable.Rows[i]["吊牌价"].ToString().ToUpper().Trim()) >= 90)
                                                codePrice = 1;
                                        }
                                        else if (Convert.ToDecimal(dtTable.Rows[i]["吊牌价"].ToString().ToUpper().Trim()) < Convert.ToDecimal(dtBiaozhunBrand.Rows[a]["吊牌价"].ToString().ToUpper().Trim()))
                                        {
                                            if (Convert.ToDecimal(dtTable.Rows[i]["吊牌价"].ToString().ToUpper().Trim()) * 100 / Convert.ToDecimal(dtBiaozhunBrand.Rows[a]["吊牌价"].ToString().ToUpper().Trim()) >= 90)
                                                codePrice = 1;
                                        }
                                    }
                                }
                                //最终匹配度
                                decimal endCode = Convert.ToDecimal(0.25) * (codePrice + codeHigh + codeClose + codeSex);
                                if (endCode >= Convert.ToDecimal(0.75))//匹配度>=0.75 则 判断（品类名称）是否相等
                                {
                                    if (dtTable.Rows[i]["品类名称"].ToString().ToUpper().Trim() == dtBiaozhunBrand.Rows[a]["品类名称"].ToString().ToUpper().Trim())//第三步（品类名称相等）
                                    {
                                        need = 0;
                                        dtBiaozhunData.Rows.Add(dtTable.Rows[i].ItemArray);
                                        dtBiaozhunData.Rows[dtBiaozhunData.Rows.Count - 1]["avc商品id"] = dtBiaozhunBrand.Rows[a]["avc商品id"];
                                        dtBiaozhunData.Rows[dtBiaozhunData.Rows.Count - 1]["avc品牌id"] = dtBiaozhunBrand.Rows[a]["avc品牌id"];
                                        break;//跳出查找循环
                                    }
                                }
                            }
                            else//品类名称不相等 属性匹配度=1
                            {
                                //开始计算 属性匹配度                                            
                                //属性相同=1，不同=0，NULL=空
                                //得分=0.25*吊牌价匹配+0.25*鞋帮高度匹配+0.25*闭合方式+0.25*性别
                                decimal codePrice = 0, codeHigh = 0, codeClose = 0, codeSex = 0;
                                if (dtTable.Rows[i]["鞋帮高度"].ToString().ToUpper().Trim() == dtBiaozhunBrand.Rows[a]["鞋帮高度"].ToString().ToUpper().Trim())
                                    codeHigh = 1;
                                if (dtTable.Rows[i]["闭合方式"].ToString().ToUpper().Trim() == dtBiaozhunBrand.Rows[a]["闭合方式"].ToString().ToUpper().Trim())
                                    codeClose = 1;
                                if (dtTable.Rows[i]["性别"].ToString().ToUpper().Trim() == dtBiaozhunBrand.Rows[a]["性别"].ToString().ToUpper().Trim())
                                    codeSex = 1;
                                if (dtTable.Rows[i]["吊牌价"].ToString().ToUpper().Trim() == dtBiaozhunBrand.Rows[a]["吊牌价"].ToString().ToUpper().Trim())
                                    codePrice = 1;
                                if (dtTable.Rows[i]["吊牌价"].ToString().ToUpper().Trim() != "NULL" && dtTable.Rows[i]["吊牌价"].ToString().ToUpper().Trim() != "")
                                {
                                    if (dtBiaozhunBrand.Rows[a]["吊牌价"].ToString().ToUpper().Trim() != "NULL" && dtBiaozhunBrand.Rows[a]["吊牌价"].ToString().ToUpper().Trim() != "")
                                    {
                                        if (Convert.ToDecimal(dtTable.Rows[i]["吊牌价"].ToString().ToUpper().Trim()) > Convert.ToDecimal(dtBiaozhunBrand.Rows[a]["吊牌价"].ToString().ToUpper().Trim()))
                                        {
                                            if (Convert.ToDecimal(dtBiaozhunBrand.Rows[a]["吊牌价"].ToString().ToUpper().Trim()) * 100 / Convert.ToDecimal(dtTable.Rows[i]["吊牌价"].ToString().ToUpper().Trim()) >= 90)
                                                codePrice = 1;
                                        }
                                        else if (Convert.ToDecimal(dtTable.Rows[i]["吊牌价"].ToString().ToUpper().Trim()) < Convert.ToDecimal(dtBiaozhunBrand.Rows[a]["吊牌价"].ToString().ToUpper().Trim()))
                                        {
                                            if (Convert.ToDecimal(dtTable.Rows[i]["吊牌价"].ToString().ToUpper().Trim()) * 100 / Convert.ToDecimal(dtBiaozhunBrand.Rows[a]["吊牌价"].ToString().ToUpper().Trim()) >= 90)
                                                codePrice = 1;
                                        }
                                    }
                                }
                                //最终匹配度
                                decimal endCode = Convert.ToDecimal(0.25) * (codePrice + codeHigh + codeClose + codeSex);
                                if (endCode == Convert.ToDecimal(1))//匹配度==1 则 判断（品类名称）是否相等
                                {
                                    if (dtTable.Rows[i]["品类名称"].ToString().ToUpper().Trim() == dtBiaozhunBrand.Rows[a]["品类名称"].ToString().ToUpper().Trim())//第三步（品类名称相等）
                                    {
                                        need = 0;
                                        dtBiaozhunData.Rows.Add(dtTable.Rows[i].ItemArray);
                                        dtBiaozhunData.Rows[dtBiaozhunData.Rows.Count - 1]["avc商品id"] = dtBiaozhunBrand.Rows[a]["avc商品id"];
                                        dtBiaozhunData.Rows[dtBiaozhunData.Rows.Count - 1]["avc品牌id"] = dtBiaozhunBrand.Rows[a]["avc品牌id"];
                                        break;//跳出查找循环
                                    }
                                }
                            }
                        }
                    }

                    #endregion

                    #endregion
                }
                if (need == 1)
                {
                    var que = (from p in dtBiaozhunBrand.AsEnumerable() group p by Convert.ToString(p.Field<object>("avc商品id")) into g orderby Convert.ToInt32(g.Key.Replace("avc00_", "")) descending select g.Key).Take(1);
                    int add = 0;
                    foreach (var q in que)
                    {
                        add = Convert.ToInt32(q.Replace("avc00_", "")) + 1;
                    }
                    dtBiaozhunBrand.Rows.Add(dtTable.Rows[i].ItemArray);//添加为新标准品牌
                    dtBiaozhunBrand.Rows[dtBiaozhunBrand.Rows.Count - 1]["avc商品id"] = "avc00_" + add + "";
                    dtBiaozhunBrand.Rows[dtBiaozhunBrand.Rows.Count - 1]["avc品牌id"] = "avcb00_" + add + "";
                    //添加为标准数据
                    dtBiaozhunData.Rows.Add(dtTable.Rows[i].ItemArray);
                    dtBiaozhunData.Rows[dtBiaozhunData.Rows.Count - 1]["avc商品id"] = "avc00_" + add + "";
                    dtBiaozhunData.Rows[dtBiaozhunData.Rows.Count - 1]["avc品牌id"] = "avcb00_" + add + "";
                }
                #endregion
                this.Invoke(new ThreadStart(delegate()
                {
                    ProgressBarDisplay(dtTable, i);
                }));
            }
            //记录时间
            File.AppendAllText(@"timeText.txt", "【sheet--" + tableName + "】开始时间：" + strDatetime + " ----结束时间：" + DateTime.Now.ToString() + "\r\n");
            //导出数据
            //mySql.DataExportToFile(dtBiaozhunData, @"avc" + tableName + ".xls");

            #endregion
        }

        /// <summary>
        /// 字符串匹配近似度
        /// </summary>
        /// <param name="sourceStr1">字符1</param>
        /// <param name="sourceStr2">字符2</param>
        /// <returns></returns>
        private int Approximate(string sourceStr1, string sourceStr2)
        {
            if (sourceStr1 == "" || sourceStr2 == "")
                return -1;
            int mark = 0;
            char[] schars1 = sourceStr1.ToCharArray();
            char[] schars2 = sourceStr2.ToCharArray();
            if (schars1.Length >= schars2.Length)
            {
                for (int i = 0; i < schars2.Length; i++)
                {
                    if (schars1[i] == schars2[i])
                        mark += 1;
                    else
                        break;
                }
                if (mark == schars2.Length)
                    return 100;
                else
                {
                    return (mark * 100 / schars2.Length);
                }
            }
            else
            {
                for (int i = 0; i < schars1.Length; i++)
                {
                    if (schars1[i] == schars2[i])
                        mark += 1;
                }
                if (mark == schars1.Length)
                    return 100;
                else
                {
                    return (mark * 100 / schars1.Length);
                }
            }
            return 0;
        }

        /// <summary>
        /// 程序优化方法
        /// </summary>
        /// <param name="dtTable"></param>
        /// <param name="tableName"></param>
        private void TheThirdMethod(DataTable dtTable, string tableName)
        {
            //商品id、商品名称、店铺名称、品牌id、品牌名称、品类id、品类名称、吊牌价、运动鞋分类 
            #region 测试（复杂方法）
            DataTable dtBiaozhunBrand = new DataTable();//存放标准品牌
            DataTable dtBiaozhunData = new DataTable();//存放查找到的标准数据
            //加载列名
            for (int i = 0; i < dtTable.Columns.Count; i++)
            {
                dtBiaozhunBrand.Columns.Add(dtTable.Columns[i].ColumnName);
                dtBiaozhunData.Columns.Add(dtTable.Columns[i].ColumnName);
            }
            //dtBiaozhunBrand、dtBiaozhunData在添加一个 标准字段，avc自定义（avc商品id、avc品牌id）
            dtBiaozhunBrand.Columns.Add("avc商品id", typeof(string));
            dtBiaozhunBrand.Columns.Add("avc品牌id", typeof(string));
            dtBiaozhunData.Columns.Add("avc商品id", typeof(string));
            dtBiaozhunData.Columns.Add("avc品牌id", typeof(string));

            for (int i = 0; i < dtTable.Rows.Count; i++)
            {
                //if (i == 0 && !specialBrands.Contains(dtTable.Rows[i]["品牌名称"].ToString().ToUpper().Trim()))//第一条数据默认为标准数据
                if (i == 0)//第一条数据默认为标准数据
                {
                    dtBiaozhunBrand.Rows.Add(dtTable.Rows[i].ItemArray);//直接复制数据 行 
                    dtBiaozhunBrand.Rows[dtBiaozhunBrand.Rows.Count - 1]["avc商品id"] = "avc00_" + i + "";
                    dtBiaozhunBrand.Rows[dtBiaozhunBrand.Rows.Count - 1]["avc品牌id"] = "avcb00_" + i + "";
                    dtBiaozhunData.Rows.Add(dtTable.Rows[i].ItemArray);
                    dtBiaozhunData.Rows[dtBiaozhunData.Rows.Count - 1]["avc商品id"] = "avc00_" + i + "";
                    dtBiaozhunData.Rows[dtBiaozhunData.Rows.Count - 1]["avc品牌id"] = "avcb00_" + i + "";
                    continue;//继续下条数据
                }

                //开始逐条数据查找判断
                #region （从第二条数据开始） 开始逐条数据查找判断
                int need = 1;//查看该条数据是否在标准列表里面查找到.
                #region 非特殊品牌处理
                if (!specialBrands.Contains(dtTable.Rows[i]["品牌名称"].ToString().ToUpper().Trim()))
                {
                    var brandId = (from p in dtBiaozhunBrand.Select("品牌名称='" + dtTable.Rows[i]["品牌名称"] + "' AND 吊牌价='" + dtTable.Rows[i]["吊牌价"].ToString().ToUpper().Trim() + "'").AsParallel() group p by new { dpj = Convert.ToString(p.Field<object>("吊牌价")), avcspid = Convert.ToString(p.Field<object>("avc商品id")), avcbrandid = Convert.ToString(p.Field<object>("avc品牌id")), kh = Convert.ToString(p.Field<object>("款号")) } into g select g.Key);
                    if (brandId.Count() > 0)
                    {
                        foreach (var q in brandId)
                        {
                            if (Approximate(dtTable.Rows[i]["款号"].ToString().ToUpper().Trim(), q.kh.ToUpper().Trim()) >= 91)
                            {
                                need = 0;
                                dtBiaozhunData.Rows.Add(dtTable.Rows[i].ItemArray);
                                dtBiaozhunData.Rows[dtBiaozhunData.Rows.Count - 1]["avc商品id"] = q.avcspid;
                                dtBiaozhunData.Rows[dtBiaozhunData.Rows.Count - 1]["avc品牌id"] = q.avcbrandid;
                                break;//跳出查找循环
                            }
                        }
                    }
                }
                #endregion

                #region 特殊品牌处理

                string[] GirardData = null;//原始数据截取
                string[] GirardBrand = null;//标准数据截取
                #region 款号只保留-之前的字符串
                if (specialBrands2.Contains(dtTable.Rows[i]["品牌名称"].ToString().ToUpper().Trim()))//款号只保留-之前的字符串
                {
                    //款号截取
                    GirardData = Regex.Split(dtTable.Rows[i]["款号"].ToString().ToUpper().Trim(), "-");
                    var queBrand = from p in dtBiaozhunBrand.Select("品牌名称='" + dtTable.Rows[i]["品牌名称"].ToString().ToUpper().Trim() + "' ").AsParallel()
                                   group p by new { xbgd = Convert.ToString(p.Field<object>("鞋帮高度")), bhfs = Convert.ToString(p.Field<object>("闭合方式")), xb = Convert.ToString(p.Field<object>("性别")), dpj = Convert.ToString(p.Field<object>("吊牌价")), plid = Convert.ToString(p.Field<object>("品类id")), avcplid = Convert.ToString(p.Field<object>("avc商品id")), avcbrandid = Convert.ToString(p.Field<object>("avc品牌id")), kh = Convert.ToString(p.Field<object>("款号")) } into g
                                   select g;
                    if (queBrand.Count() > 0)
                    {
                        foreach (var q in queBrand)
                        {
                            GirardBrand = Regex.Split(q.Key.kh.ToUpper().Trim(), "-");

                            if (GirardData[0] == GirardBrand[0])
                            {
                                //开始计算 属性匹配度                                            
                                //属性相同=1，不同=0，NULL=空
                                //得分=0.25*吊牌价匹配+0.25*鞋帮高度匹配+0.25*闭合方式+0.25*性别
                                decimal codePrice = 0, codeHigh = 0, codeClose = 0, codeSex = 0;
                                if (dtTable.Rows[i]["鞋帮高度"].ToString().ToUpper().Trim() == q.Key.xbgd.ToUpper().Trim())
                                    codeHigh = 1;
                                if (dtTable.Rows[i]["闭合方式"].ToString().ToUpper().Trim() == q.Key.bhfs.ToUpper().Trim())
                                    codeClose = 1;
                                if (dtTable.Rows[i]["性别"].ToString().ToUpper().Trim() == q.Key.xb.ToUpper().Trim())
                                    codeSex = 1;
                                if (q.Key.dpj.ToUpper() != "NULL" && q.Key.dpj.ToUpper() != "" && dtTable.Rows[i]["吊牌价"].ToString().ToUpper().Trim() != "NULL" && dtTable.Rows[i]["吊牌价"].ToString().ToUpper().Trim() != "")
                                {
                                    if (Convert.ToDecimal(dtTable.Rows[i]["吊牌价"].ToString().ToUpper().Trim()) >= Convert.ToDecimal(q.Key.dpj))
                                    {
                                        if (Convert.ToDecimal(q.Key.dpj) * 100 / Convert.ToDecimal(dtTable.Rows[i]["吊牌价"].ToString().ToUpper().Trim()) >= 90)
                                            codePrice = 1;
                                    }
                                    else
                                    {
                                        if (Convert.ToDecimal(dtTable.Rows[i]["吊牌价"].ToString().ToUpper().Trim()) * 100 / Convert.ToDecimal(q.Key.dpj) >= 90)
                                            codePrice = 1;
                                    }
                                }
                                //最终匹配度
                                decimal endCode = Convert.ToDecimal(0.25) * (codePrice + codeHigh + codeClose + codeSex);
                                if (dtTable.Rows[i]["品类id"].ToString().ToUpper().Trim() == q.Key.plid.ToUpper().Trim())
                                {
                                    if (endCode >= Convert.ToDecimal(0.75))//匹配度>=0.75 则 判断（品类名称）是否相等
                                    {
                                        need = 0;
                                        dtBiaozhunData.Rows.Add(dtTable.Rows[i].ItemArray);
                                        dtBiaozhunData.Rows[dtBiaozhunData.Rows.Count - 1]["avc商品id"] = q.Key.avcplid;
                                        dtBiaozhunData.Rows[dtBiaozhunData.Rows.Count - 1]["avc品牌id"] = q.Key.avcbrandid;
                                        break;//跳出查找循环
                                    }
                                }
                                else
                                {
                                    if (endCode == Convert.ToDecimal(1))//匹配度>=0.75 则 判断（品类名称）是否相等
                                    {
                                        need = 0;
                                        dtBiaozhunData.Rows.Add(dtTable.Rows[i].ItemArray);
                                        dtBiaozhunData.Rows[dtBiaozhunData.Rows.Count - 1]["avc商品id"] = q.Key.avcplid;
                                        dtBiaozhunData.Rows[dtBiaozhunData.Rows.Count - 1]["avc品牌id"] = q.Key.avcbrandid;
                                        break;//跳出查找循环
                                    }
                                }
                            }
                        }
                    }
                    ////标准款号截取
                    //GirardBrand = Regex.Split(dtBiaozhunBrand.Rows[a]["款号"].ToString().ToUpper().Trim(), "-");
                }
                #endregion
                #region 将空格*/等去除，保留前六位数字（Adidas/阿迪达斯：2015SSOR-KCO11 款式格式为此，不参与）
                else if (specialBrands3.Contains(dtTable.Rows[i]["品牌名称"].ToString().ToUpper().Trim()))//将空格*/等去除，保留前六位数字（Adidas/阿迪达斯：2015SSOR-KCO11 款式格式为此，不参与）
                {
                    var queBrand = from p in dtBiaozhunBrand.Select("品牌名称='" + dtTable.Rows[i]["品牌名称"].ToString().ToUpper().Trim() + "' AND 款号 LIKE '%-%' AND 款号 LIKE '201%'").AsParallel() group p by new { xbgd = Convert.ToString(p.Field<object>("鞋帮高度")), bhfs = Convert.ToString(p.Field<object>("闭合方式")), xb = Convert.ToString(p.Field<object>("性别")), dpj = Convert.ToString(p.Field<object>("吊牌价")), plid = Convert.ToString(p.Field<object>("品类id")), avcplid = Convert.ToString(p.Field<object>("avc商品id")), avcbrandid = Convert.ToString(p.Field<object>("avc品牌id")), kh = Convert.ToString(p.Field<object>("款号")) } into g select g;
                    if (queBrand.Count() > 0)
                    {
                        foreach (var q in queBrand)
                        {
                            if (q.Key.kh.ToUpper().Trim() == dtTable.Rows[i]["款号"].ToString().ToUpper().Trim())//判断款号是否相等
                            {
                                //开始计算 属性匹配度                                            
                                //属性相同=1，不同=0，NULL=空
                                //得分=0.25*吊牌价匹配+0.25*鞋帮高度匹配+0.25*闭合方式+0.25*性别
                                decimal codePrice = 0, codeHigh = 0, codeClose = 0, codeSex = 0;
                                if (dtTable.Rows[i]["鞋帮高度"].ToString().ToUpper().Trim() == q.Key.xbgd.ToUpper().Trim())
                                    codeHigh = 1;
                                if (dtTable.Rows[i]["闭合方式"].ToString().ToUpper().Trim() == q.Key.bhfs.ToUpper().Trim())
                                    codeClose = 1;
                                if (dtTable.Rows[i]["性别"].ToString().ToUpper().Trim() == q.Key.xb.ToUpper().Trim())
                                    codeSex = 1;
                                if (q.Key.dpj.ToUpper() != "NULL" && q.Key.dpj.ToUpper() != "" && dtTable.Rows[i]["吊牌价"].ToString().ToUpper().Trim() != "NULL" && dtTable.Rows[i]["吊牌价"].ToString().ToUpper().Trim() != "")
                                {
                                    if (Convert.ToDecimal(dtTable.Rows[i]["吊牌价"].ToString().ToUpper().Trim()) >= Convert.ToDecimal(q.Key.dpj))
                                    {
                                        if (Convert.ToDecimal(q.Key.dpj) * 100 / Convert.ToDecimal(dtTable.Rows[i]["吊牌价"].ToString().ToUpper().Trim()) >= 90)
                                            codePrice = 1;
                                    }
                                    else
                                    {
                                        if (Convert.ToDecimal(dtTable.Rows[i]["吊牌价"].ToString().ToUpper().Trim()) * 100 / Convert.ToDecimal(q.Key.dpj) >= 90)
                                            codePrice = 1;
                                    }
                                }
                                //最终匹配度
                                decimal endCode = Convert.ToDecimal(0.25) * (codePrice + codeHigh + codeClose + codeSex);
                                if (dtTable.Rows[i]["品类id"].ToString().ToUpper().Trim() == q.Key.plid.ToUpper().Trim())
                                {
                                    if (endCode >= Convert.ToDecimal(0.75))//匹配度>=0.75 则 判断（品类名称相等）是否相等
                                    {
                                        need = 0;
                                        dtBiaozhunData.Rows.Add(dtTable.Rows[i].ItemArray);
                                        dtBiaozhunData.Rows[dtBiaozhunData.Rows.Count - 1]["avc商品id"] = q.Key.avcplid;
                                        dtBiaozhunData.Rows[dtBiaozhunData.Rows.Count - 1]["avc品牌id"] = q.Key.avcbrandid;
                                        break;//跳出查找循环
                                    }
                                }
                                else
                                {
                                    if (endCode == Convert.ToDecimal(1))//匹配度==1 则 判断（品类名称不相等）是否相等
                                    {
                                        need = 0;
                                        dtBiaozhunData.Rows.Add(dtTable.Rows[i].ItemArray);
                                        dtBiaozhunData.Rows[dtBiaozhunData.Rows.Count - 1]["avc商品id"] = q.Key.avcplid;
                                        dtBiaozhunData.Rows[dtBiaozhunData.Rows.Count - 1]["avc品牌id"] = q.Key.avcbrandid;
                                        break;//跳出查找循环
                                    }
                                }
                            }
                            else//款号不相等,则为新标准品牌
                            {
                                need = 0;
                                var que = (from p in dtBiaozhunBrand.AsEnumerable() group p by Convert.ToString(p.Field<object>("avc商品id")) into g orderby Convert.ToInt32(g.Key.Replace("avc00_", "")) descending select g.Key).Take(1);
                                int add = 0;
                                foreach (var qx in que)
                                {
                                    add = Convert.ToInt32(qx.Replace("avc00_", "")) + 1;
                                }
                                dtBiaozhunBrand.Rows.Add(dtTable.Rows[i].ItemArray);//添加为新标准品牌
                                dtBiaozhunBrand.Rows[dtBiaozhunBrand.Rows.Count - 1]["avc商品id"] = "avc00_" + add + "";
                                dtBiaozhunBrand.Rows[dtBiaozhunBrand.Rows.Count - 1]["avc品牌id"] = "avcb00_" + add + "";
                                //添加为标准数据
                                dtBiaozhunData.Rows.Add(dtTable.Rows[i].ItemArray);
                                dtBiaozhunData.Rows[dtBiaozhunData.Rows.Count - 1]["avc商品id"] = "avc00_" + add + "";
                                dtBiaozhunData.Rows[dtBiaozhunData.Rows.Count - 1]["avc品牌id"] = "avcb00_" + add + "";
                                break;
                            }
                        }
                    }
                    else
                    {
                        string girardData1 = string.Empty;
                        string girardData2 = string.Empty;
                        girardData1 = Regex.Replace(dtTable.Rows[i]["款号"].ToString().ToUpper().Trim(), @"[^a-zA-Z0-9\u4E00-\u9FA5\uF900-\uFA2D]", "").Replace(" ", "");
                        if (girardData1.Length > 6)
                            girardData1 = girardData1.Substring(0, 6);
                        //款号截取
                        GirardData = Regex.Split(girardData1, "~");

                        var queBrandkh = from p in dtBiaozhunBrand.Select("品牌名称='" + dtTable.Rows[i]["品牌名称"].ToString().ToUpper().Trim() + "'") group p by new { xbgd = Convert.ToString(p.Field<object>("鞋帮高度")), bhfs = Convert.ToString(p.Field<object>("闭合方式")), xb = Convert.ToString(p.Field<object>("性别")), dpj = Convert.ToString(p.Field<object>("吊牌价")), plid = Convert.ToString(p.Field<object>("品类id")), avcplid = Convert.ToString(p.Field<object>("avc商品id")), avcbrandid = Convert.ToString(p.Field<object>("avc品牌id")), kh = Convert.ToString(p.Field<object>("款号")) } into g select g;
                        foreach (var q in queBrandkh)
                        {
                            girardData2 = Regex.Replace(q.Key.kh.ToString().ToUpper().Trim(), @"[^a-zA-Z0-9\u4E00-\u9FA5\uF900-\uFA2D]", "").Replace(" ", "");
                            if (girardData2.Length > 6)
                                girardData2 = girardData2.Substring(0, 6);
                            GirardBrand = Regex.Split(girardData2, "~");
                            if (GirardData[0] == GirardBrand[0])
                            {
                                //开始计算 属性匹配度                                            
                                //属性相同=1，不同=0，NULL=空
                                //得分=0.25*吊牌价匹配+0.25*鞋帮高度匹配+0.25*闭合方式+0.25*性别
                                decimal codePrice = 0, codeHigh = 0, codeClose = 0, codeSex = 0;
                                if (dtTable.Rows[i]["鞋帮高度"].ToString().ToUpper().Trim() == q.Key.xbgd.ToUpper().Trim())
                                    codeHigh = 1;
                                if (dtTable.Rows[i]["闭合方式"].ToString().ToUpper().Trim() == q.Key.bhfs.ToUpper().Trim())
                                    codeClose = 1;
                                if (dtTable.Rows[i]["性别"].ToString().ToUpper().Trim() == q.Key.xb.ToUpper().Trim())
                                    codeSex = 1;
                                if (q.Key.dpj.ToUpper() != "NULL" && q.Key.dpj.ToUpper() != "" && dtTable.Rows[i]["吊牌价"].ToString().ToUpper().Trim() != "NULL" && dtTable.Rows[i]["吊牌价"].ToString().ToUpper().Trim() != "")
                                {
                                    if (Convert.ToDecimal(dtTable.Rows[i]["吊牌价"].ToString().ToUpper().Trim()) >= Convert.ToDecimal(q.Key.dpj))
                                    {
                                        if (Convert.ToDecimal(q.Key.dpj) * 100 / Convert.ToDecimal(dtTable.Rows[i]["吊牌价"].ToString().ToUpper().Trim()) >= 90)
                                            codePrice = 1;
                                    }
                                    else
                                    {
                                        if (Convert.ToDecimal(dtTable.Rows[i]["吊牌价"].ToString().ToUpper().Trim()) * 100 / Convert.ToDecimal(q.Key.dpj) >= 90)
                                            codePrice = 1;
                                    }
                                }
                                //最终匹配度
                                decimal endCode = Convert.ToDecimal(0.25) * (codePrice + codeHigh + codeClose + codeSex);
                                if (dtTable.Rows[i]["品类id"].ToString().ToUpper().Trim() == q.Key.plid.ToUpper().Trim())
                                {
                                    if (endCode >= Convert.ToDecimal(0.75))//匹配度>=0.75 则 判断（品类名称）是否相等
                                    {
                                        need = 0;
                                        dtBiaozhunData.Rows.Add(dtTable.Rows[i].ItemArray);
                                        dtBiaozhunData.Rows[dtBiaozhunData.Rows.Count - 1]["avc商品id"] = q.Key.avcplid;
                                        dtBiaozhunData.Rows[dtBiaozhunData.Rows.Count - 1]["avc品牌id"] = q.Key.avcbrandid;
                                        break;//跳出查找循环
                                    }
                                }
                                else
                                {
                                    if (endCode == Convert.ToDecimal(1))//匹配度>=0.75 则 判断（品类名称）是否相等
                                    {
                                        need = 0;
                                        dtBiaozhunData.Rows.Add(dtTable.Rows[i].ItemArray);
                                        dtBiaozhunData.Rows[dtBiaozhunData.Rows.Count - 1]["avc商品id"] = q.Key.avcplid;
                                        dtBiaozhunData.Rows[dtBiaozhunData.Rows.Count - 1]["avc品牌id"] = q.Key.avcbrandid;
                                        break;//跳出查找循环
                                    }
                                }
                            }
                        }
                    }
                }
                #endregion
                #region 去除字母只保留数字
                else if (specialBrands4.Contains(dtTable.Rows[i]["品牌名称"].ToString().ToUpper().Trim()))//款号只保留-之前的字符串
                {
                    string girardData1 = string.Empty;
                    string girardData2 = string.Empty;
                    Regex regex = new Regex(@"[0-9]+");//所有数字
                    girardData1 = regex.Match(dtTable.Rows[i]["款号"].ToString().ToUpper().Trim()).ToString();
                    //girardData2 = regex.Match(dtBiaozhunBrand.Rows[a]["款号"].ToString().ToUpper().Trim()).ToString();
                    //款号截取
                    GirardData = Regex.Split(girardData1, "~");
                    //标准款号截取
                    //GirardBrand = Regex.Split(girardData2, "~");
                    var queBrand = from p in dtBiaozhunBrand.Select("品牌名称='" + dtTable.Rows[i]["品牌名称"].ToString().ToUpper().Trim() + "' ").AsParallel()
                                   group p by new { xbgd = Convert.ToString(p.Field<object>("鞋帮高度")), bhfs = Convert.ToString(p.Field<object>("闭合方式")), xb = Convert.ToString(p.Field<object>("性别")), dpj = Convert.ToString(p.Field<object>("吊牌价")), plid = Convert.ToString(p.Field<object>("品类id")), avcplid = Convert.ToString(p.Field<object>("avc商品id")), avcbrandid = Convert.ToString(p.Field<object>("avc品牌id")), kh = Convert.ToString(p.Field<object>("款号")) } into g
                                   select g;
                    if (queBrand.Count() > 0)
                    {
                        foreach (var q in queBrand)
                        {
                            //标准款号截取
                            girardData2 = regex.Match(q.Key.kh.ToUpper().Trim()).ToString();
                            GirardBrand = Regex.Split(girardData2, "~");

                            if (GirardData[0] == GirardBrand[0])
                            {
                                //开始计算 属性匹配度                                            
                                //属性相同=1，不同=0，NULL=空
                                //得分=0.25*吊牌价匹配+0.25*鞋帮高度匹配+0.25*闭合方式+0.25*性别
                                decimal codePrice = 0, codeHigh = 0, codeClose = 0, codeSex = 0;
                                if (dtTable.Rows[i]["鞋帮高度"].ToString().ToUpper().Trim() == q.Key.xbgd.ToUpper().Trim())
                                    codeHigh = 1;
                                if (dtTable.Rows[i]["闭合方式"].ToString().ToUpper().Trim() == q.Key.bhfs.ToUpper().Trim())
                                    codeClose = 1;
                                if (dtTable.Rows[i]["性别"].ToString().ToUpper().Trim() == q.Key.xb.ToUpper().Trim())
                                    codeSex = 1;
                                if (q.Key.dpj.ToUpper() != "NULL" && q.Key.dpj.ToUpper() != "" && dtTable.Rows[i]["吊牌价"].ToString().ToUpper().Trim() != "NULL" && dtTable.Rows[i]["吊牌价"].ToString().ToUpper().Trim() != "")
                                {
                                    if (Convert.ToDecimal(dtTable.Rows[i]["吊牌价"].ToString().ToUpper().Trim()) >= Convert.ToDecimal(q.Key.dpj))
                                    {
                                        if (Convert.ToDecimal(q.Key.dpj) * 100 / Convert.ToDecimal(dtTable.Rows[i]["吊牌价"].ToString().ToUpper().Trim()) >= 90)
                                            codePrice = 1;
                                    }
                                    else
                                    {
                                        if (Convert.ToDecimal(dtTable.Rows[i]["吊牌价"].ToString().ToUpper().Trim()) * 100 / Convert.ToDecimal(q.Key.dpj) >= 90)
                                            codePrice = 1;
                                    }
                                }
                                //最终匹配度
                                decimal endCode = Convert.ToDecimal(0.25) * (codePrice + codeHigh + codeClose + codeSex);
                                if (dtTable.Rows[i]["品类id"].ToString().ToUpper().Trim() == q.Key.plid.ToUpper().Trim())
                                {
                                    if (endCode >= Convert.ToDecimal(0.75))//匹配度>=0.75 则 判断（品类名称）是否相等
                                    {
                                        need = 0;
                                        dtBiaozhunData.Rows.Add(dtTable.Rows[i].ItemArray);
                                        dtBiaozhunData.Rows[dtBiaozhunData.Rows.Count - 1]["avc商品id"] = q.Key.avcplid;
                                        dtBiaozhunData.Rows[dtBiaozhunData.Rows.Count - 1]["avc品牌id"] = q.Key.avcbrandid;
                                        break;//跳出查找循环
                                    }
                                }
                                else
                                {
                                    if (endCode == Convert.ToDecimal(1))//匹配度>=0.75 则 判断（品类名称）是否相等
                                    {
                                        need = 0;
                                        dtBiaozhunData.Rows.Add(dtTable.Rows[i].ItemArray);
                                        dtBiaozhunData.Rows[dtBiaozhunData.Rows.Count - 1]["avc商品id"] = q.Key.avcplid;
                                        dtBiaozhunData.Rows[dtBiaozhunData.Rows.Count - 1]["avc品牌id"] = q.Key.avcbrandid;
                                        break;//跳出查找循环
                                    }
                                }
                            }
                        }
                    }
                }
                #endregion
                #region 去除字母只保留数字，且只保留-之前的字符串
                else if (specialBrands5.Contains(dtTable.Rows[i]["品牌名称"].ToString().ToUpper().Trim()))//去除字母只保留数字，且只保留-之前的字符串
                {
                    string girardData1 = string.Empty;
                    string girardData2 = string.Empty;
                    Regex regex = new Regex(@"[0-9]+");//所有数字
                    girardData1 = regex.Match(dtTable.Rows[i]["款号"].ToString().ToUpper().Trim()).ToString();
                    //girardData2 = regex.Match(dtBiaozhunBrand.Rows[a]["款号"].ToString().ToUpper().Trim()).ToString();
                    GirardData = Regex.Split(girardData1, "-");
                    //GirardBrand = Regex.Split(girardData2, "-");
                    var queBrand = from p in dtBiaozhunBrand.Select("品牌名称='" + dtTable.Rows[i]["品牌名称"].ToString().ToUpper().Trim() + "' ").AsParallel()
                                   group p by new { xbgd = Convert.ToString(p.Field<object>("鞋帮高度")), bhfs = Convert.ToString(p.Field<object>("闭合方式")), xb = Convert.ToString(p.Field<object>("性别")), dpj = Convert.ToString(p.Field<object>("吊牌价")), plid = Convert.ToString(p.Field<object>("品类id")), avcplid = Convert.ToString(p.Field<object>("avc商品id")), avcbrandid = Convert.ToString(p.Field<object>("avc品牌id")), kh = Convert.ToString(p.Field<object>("款号")) } into g
                                   select g;
                    if (queBrand.Count() > 0)
                    {
                        foreach (var q in queBrand)
                        {
                            //标准款号截取
                            girardData2 = regex.Match(q.Key.kh.ToUpper().Trim()).ToString();
                            GirardBrand = Regex.Split(girardData2, "-");

                            if (GirardData[0] == GirardBrand[0])
                            {
                                //开始计算 属性匹配度                                            
                                //属性相同=1，不同=0，NULL=空
                                //得分=0.25*吊牌价匹配+0.25*鞋帮高度匹配+0.25*闭合方式+0.25*性别
                                decimal codePrice = 0, codeHigh = 0, codeClose = 0, codeSex = 0;
                                if (dtTable.Rows[i]["鞋帮高度"].ToString().ToUpper().Trim() == q.Key.xbgd.ToUpper().Trim())
                                    codeHigh = 1;
                                if (dtTable.Rows[i]["闭合方式"].ToString().ToUpper().Trim() == q.Key.bhfs.ToUpper().Trim())
                                    codeClose = 1;
                                if (dtTable.Rows[i]["性别"].ToString().ToUpper().Trim() == q.Key.xb.ToUpper().Trim())
                                    codeSex = 1;
                                if (q.Key.dpj.ToUpper() != "NULL" && q.Key.dpj.ToUpper() != "" && dtTable.Rows[i]["吊牌价"].ToString().ToUpper().Trim() != "NULL" && dtTable.Rows[i]["吊牌价"].ToString().ToUpper().Trim() != "")
                                {
                                    if (Convert.ToDecimal(dtTable.Rows[i]["吊牌价"].ToString().ToUpper().Trim()) >= Convert.ToDecimal(q.Key.dpj))
                                    {
                                        if (Convert.ToDecimal(q.Key.dpj) * 100 / Convert.ToDecimal(dtTable.Rows[i]["吊牌价"].ToString().ToUpper().Trim()) >= 90)
                                            codePrice = 1;
                                    }
                                    else
                                    {
                                        if (Convert.ToDecimal(dtTable.Rows[i]["吊牌价"].ToString().ToUpper().Trim()) * 100 / Convert.ToDecimal(q.Key.dpj) >= 90)
                                            codePrice = 1;
                                    }
                                }
                                //最终匹配度
                                decimal endCode = Convert.ToDecimal(0.25) * (codePrice + codeHigh + codeClose + codeSex);
                                if (dtTable.Rows[i]["品类id"].ToString().ToUpper().Trim() == q.Key.plid.ToUpper().Trim())
                                {
                                    if (endCode >= Convert.ToDecimal(0.75))//匹配度>=0.75 则 判断（品类名称）是否相等
                                    {
                                        need = 0;
                                        dtBiaozhunData.Rows.Add(dtTable.Rows[i].ItemArray);
                                        dtBiaozhunData.Rows[dtBiaozhunData.Rows.Count - 1]["avc商品id"] = q.Key.avcplid;
                                        dtBiaozhunData.Rows[dtBiaozhunData.Rows.Count - 1]["avc品牌id"] = q.Key.avcbrandid;
                                        break;//跳出查找循环
                                    }
                                }
                                else
                                {
                                    if (endCode == Convert.ToDecimal(1))//匹配度>=0.75 则 判断（品类名称）是否相等
                                    {
                                        need = 0;
                                        dtBiaozhunData.Rows.Add(dtTable.Rows[i].ItemArray);
                                        dtBiaozhunData.Rows[dtBiaozhunData.Rows.Count - 1]["avc商品id"] = q.Key.avcplid;
                                        dtBiaozhunData.Rows[dtBiaozhunData.Rows.Count - 1]["avc品牌id"] = q.Key.avcbrandid;
                                        break;//跳出查找循环
                                    }
                                }
                            }
                        }
                    }
                }
                #endregion
                #region 位如果有-需去除，保留-之前字符串
                else if (specialBrands6.Contains(dtTable.Rows[i]["品牌名称"].ToString().ToUpper().Trim()))//首位如果有-需去除，保留-之前字符串
                {
                    string girardData1 = string.Empty;
                    string girardData2 = string.Empty;
                    if (dtTable.Rows[i]["款号"].ToString().ToUpper().Trim()[0] == '-')
                        girardData1 = dtTable.Rows[i]["款号"].ToString().ToUpper().Trim().Substring(1);//清洗第一个 '-'
                    else
                        girardData1 = dtTable.Rows[i]["款号"].ToString().ToUpper().Trim();
                    //款号截取
                    GirardData = Regex.Split(girardData1, "-");
                    ////标准款号截取
                    //GirardBrand = Regex.Split(girardData2, "-");
                    var queBrand = from p in dtBiaozhunBrand.Select("品牌名称='" + dtTable.Rows[i]["品牌名称"].ToString().ToUpper().Trim() + "' ").AsParallel()
                                   group p by new { xbgd = Convert.ToString(p.Field<object>("鞋帮高度")), bhfs = Convert.ToString(p.Field<object>("闭合方式")), xb = Convert.ToString(p.Field<object>("性别")), dpj = Convert.ToString(p.Field<object>("吊牌价")), plid = Convert.ToString(p.Field<object>("品类id")), avcplid = Convert.ToString(p.Field<object>("avc商品id")), avcbrandid = Convert.ToString(p.Field<object>("avc品牌id")), kh = Convert.ToString(p.Field<object>("款号")) } into g
                                   select g;
                    if (queBrand.Count() > 0)
                    {
                        foreach (var q in queBrand)
                        {
                            //标准款号截取
                            if (q.Key.kh.ToUpper().Trim()[0] == '-')
                                girardData2 = q.Key.kh.ToUpper().Trim().Substring(1);//清洗第一个 '-'
                            else
                                girardData2 = q.Key.kh.ToUpper().Trim();
                            //标准款号截取
                            GirardBrand = Regex.Split(girardData2, "-");

                            if (GirardData[0] == GirardBrand[0])
                            {
                                //开始计算 属性匹配度                                            
                                //属性相同=1，不同=0，NULL=空
                                //得分=0.25*吊牌价匹配+0.25*鞋帮高度匹配+0.25*闭合方式+0.25*性别
                                decimal codePrice = 0, codeHigh = 0, codeClose = 0, codeSex = 0;
                                if (dtTable.Rows[i]["鞋帮高度"].ToString().ToUpper().Trim() == q.Key.xbgd.ToUpper().Trim())
                                    codeHigh = 1;
                                if (dtTable.Rows[i]["闭合方式"].ToString().ToUpper().Trim() == q.Key.bhfs.ToUpper().Trim())
                                    codeClose = 1;
                                if (dtTable.Rows[i]["性别"].ToString().ToUpper().Trim() == q.Key.xb.ToUpper().Trim())
                                    codeSex = 1;
                                if (q.Key.dpj.ToUpper() != "NULL" && q.Key.dpj.ToUpper() != "" && dtTable.Rows[i]["吊牌价"].ToString().ToUpper().Trim() != "NULL" && dtTable.Rows[i]["吊牌价"].ToString().ToUpper().Trim() != "")
                                {
                                    if (Convert.ToDecimal(dtTable.Rows[i]["吊牌价"].ToString().ToUpper().Trim()) >= Convert.ToDecimal(q.Key.dpj))
                                    {
                                        if (Convert.ToDecimal(q.Key.dpj) * 100 / Convert.ToDecimal(dtTable.Rows[i]["吊牌价"].ToString().ToUpper().Trim()) >= 90)
                                            codePrice = 1;
                                    }
                                    else
                                    {
                                        if (Convert.ToDecimal(dtTable.Rows[i]["吊牌价"].ToString().ToUpper().Trim()) * 100 / Convert.ToDecimal(q.Key.dpj) >= 90)
                                            codePrice = 1;
                                    }
                                }
                                //最终匹配度
                                decimal endCode = Convert.ToDecimal(0.25) * (codePrice + codeHigh + codeClose + codeSex);
                                if (dtTable.Rows[i]["品类id"].ToString().ToUpper().Trim() == q.Key.plid.ToUpper().Trim())
                                {
                                    if (endCode >= Convert.ToDecimal(0.75))//匹配度>=0.75 则 判断（品类名称）是否相等
                                    {
                                        need = 0;
                                        dtBiaozhunData.Rows.Add(dtTable.Rows[i].ItemArray);
                                        dtBiaozhunData.Rows[dtBiaozhunData.Rows.Count - 1]["avc商品id"] = q.Key.avcplid;
                                        dtBiaozhunData.Rows[dtBiaozhunData.Rows.Count - 1]["avc品牌id"] = q.Key.avcbrandid;
                                        break;//跳出查找循环
                                    }
                                }
                                else
                                {
                                    if (endCode == Convert.ToDecimal(1))//匹配度>=0.75 则 判断（品类名称）是否相等
                                    {
                                        need = 0;
                                        dtBiaozhunData.Rows.Add(dtTable.Rows[i].ItemArray);
                                        dtBiaozhunData.Rows[dtBiaozhunData.Rows.Count - 1]["avc商品id"] = q.Key.avcplid;
                                        dtBiaozhunData.Rows[dtBiaozhunData.Rows.Count - 1]["avc品牌id"] = q.Key.avcbrandid;
                                        break;//跳出查找循环
                                    }
                                }
                            }
                        }
                    }
                }
                #endregion
                #region 保留前六位数字
                else if (specialBrands7.Contains(dtTable.Rows[i]["品牌名称"].ToString().ToUpper().Trim()))//保留前六位数字
                {
                    string girardData1 = string.Empty;
                    string girardData2 = string.Empty;
                    Regex regex = new Regex(@"[0-9]+");//所有数字
                    girardData1 = regex.Match(dtTable.Rows[i]["款号"].ToString().ToUpper().Trim()).ToString();
                    //girardData2 = regex.Match(dtBiaozhunBrand.Rows[a]["款号"].ToString().ToUpper().Trim()).ToString();
                    //判断是否款号字符长度>6
                    if (girardData1.Length > 6)
                        girardData1 = girardData1.Substring(0, 6);
                    //if (girardData2.Length > 6)
                    //    girardData2 = girardData2.Substring(0, 6);
                    //款号截取
                    GirardData = Regex.Split(girardData1, "~");
                    ////标准款号截取
                    //GirardBrand = Regex.Split(girardData2, "~");
                    var queBrand = from p in dtBiaozhunBrand.Select("品牌名称='" + dtTable.Rows[i]["品牌名称"].ToString().ToUpper().Trim() + "' ").AsParallel()
                                   group p by new { xbgd = Convert.ToString(p.Field<object>("鞋帮高度")), bhfs = Convert.ToString(p.Field<object>("闭合方式")), xb = Convert.ToString(p.Field<object>("性别")), dpj = Convert.ToString(p.Field<object>("吊牌价")), plid = Convert.ToString(p.Field<object>("品类id")), avcplid = Convert.ToString(p.Field<object>("avc商品id")), avcbrandid = Convert.ToString(p.Field<object>("avc品牌id")), kh = Convert.ToString(p.Field<object>("款号")) } into g
                                   select g;
                    if (queBrand.Count() > 0)
                    {
                        foreach (var q in queBrand)
                        {
                            //标准款号截取
                            girardData2 = regex.Match(q.Key.kh.ToUpper().Trim()).ToString();
                            if (girardData2.Length > 6)
                                girardData2 = girardData2.Substring(0, 6);
                            //标准款号截取
                            GirardBrand = Regex.Split(girardData2, "~");

                            if (GirardData[0] == GirardBrand[0])
                            {
                                //开始计算 属性匹配度                                            
                                //属性相同=1，不同=0，NULL=空
                                //得分=0.25*吊牌价匹配+0.25*鞋帮高度匹配+0.25*闭合方式+0.25*性别
                                decimal codePrice = 0, codeHigh = 0, codeClose = 0, codeSex = 0;
                                if (dtTable.Rows[i]["鞋帮高度"].ToString().ToUpper().Trim() == q.Key.xbgd.ToUpper().Trim())
                                    codeHigh = 1;
                                if (dtTable.Rows[i]["闭合方式"].ToString().ToUpper().Trim() == q.Key.bhfs.ToUpper().Trim())
                                    codeClose = 1;
                                if (dtTable.Rows[i]["性别"].ToString().ToUpper().Trim() == q.Key.xb.ToUpper().Trim())
                                    codeSex = 1;
                                if (q.Key.dpj.ToUpper() != "NULL" && q.Key.dpj.ToUpper() != "" && dtTable.Rows[i]["吊牌价"].ToString().ToUpper().Trim() != "NULL" && dtTable.Rows[i]["吊牌价"].ToString().ToUpper().Trim() != "")
                                {
                                    if (Convert.ToDecimal(dtTable.Rows[i]["吊牌价"].ToString().ToUpper().Trim()) >= Convert.ToDecimal(q.Key.dpj))
                                    {
                                        if (Convert.ToDecimal(q.Key.dpj) * 100 / Convert.ToDecimal(dtTable.Rows[i]["吊牌价"].ToString().ToUpper().Trim()) >= 90)
                                            codePrice = 1;
                                    }
                                    else
                                    {
                                        if (Convert.ToDecimal(dtTable.Rows[i]["吊牌价"].ToString().ToUpper().Trim()) * 100 / Convert.ToDecimal(q.Key.dpj) >= 90)
                                            codePrice = 1;
                                    }
                                }
                                //最终匹配度
                                decimal endCode = Convert.ToDecimal(0.25) * (codePrice + codeHigh + codeClose + codeSex);
                                if (dtTable.Rows[i]["品类id"].ToString().ToUpper().Trim() == q.Key.plid.ToUpper().Trim())
                                {
                                    if (endCode >= Convert.ToDecimal(0.75))//匹配度>=0.75 则 判断（品类名称）是否相等
                                    {
                                        need = 0;
                                        dtBiaozhunData.Rows.Add(dtTable.Rows[i].ItemArray);
                                        dtBiaozhunData.Rows[dtBiaozhunData.Rows.Count - 1]["avc商品id"] = q.Key.avcplid;
                                        dtBiaozhunData.Rows[dtBiaozhunData.Rows.Count - 1]["avc品牌id"] = q.Key.avcbrandid;
                                        break;//跳出查找循环
                                    }
                                }
                                else
                                {
                                    if (endCode == Convert.ToDecimal(1))//匹配度>=0.75 则 判断（品类名称）是否相等
                                    {
                                        need = 0;
                                        dtBiaozhunData.Rows.Add(dtTable.Rows[i].ItemArray);
                                        dtBiaozhunData.Rows[dtBiaozhunData.Rows.Count - 1]["avc商品id"] = q.Key.avcplid;
                                        dtBiaozhunData.Rows[dtBiaozhunData.Rows.Count - 1]["avc品牌id"] = q.Key.avcbrandid;
                                        break;//跳出查找循环
                                    }
                                }
                            }
                        }
                    }
                }
                #endregion
                #region 去除尾号字母
                else if (specialBrands8.Contains(dtTable.Rows[i]["品牌名称"].ToString().ToUpper().Trim()))//去除尾号字母
                {
                    string girardData1 = string.Empty;
                    string girardData2 = string.Empty;
                    string haveabc = "";
                    Regex regex = new Regex(@"\d[a-zA-Z]");//以字母结尾
                    MatchCollection mm1 = regex.Matches(dtTable.Rows[i]["款号"].ToString().ToUpper().Trim());
                    if (mm1.Count > 0)
                    {
                        haveabc = mm1[mm1.Count - 1].Value.ToString();
                        girardData1 = dtTable.Rows[i]["款号"].ToString().ToUpper().Trim().Substring(0, dtTable.Rows[i]["款号"].ToString().ToUpper().Trim().LastIndexOf(haveabc) + 1);
                        haveabc = "";
                    }
                    else
                        girardData1 = dtTable.Rows[i]["款号"].ToString().ToUpper().Trim();
                    //款号截取
                    GirardData = Regex.Split(girardData1, "~");

                    var queBrand = from p in dtBiaozhunBrand.Select("品牌名称='" + dtTable.Rows[i]["品牌名称"].ToString().ToUpper().Trim() + "' ").AsParallel()
                                   group p by new { xbgd = Convert.ToString(p.Field<object>("鞋帮高度")), bhfs = Convert.ToString(p.Field<object>("闭合方式")), xb = Convert.ToString(p.Field<object>("性别")), dpj = Convert.ToString(p.Field<object>("吊牌价")), plid = Convert.ToString(p.Field<object>("品类id")), avcplid = Convert.ToString(p.Field<object>("avc商品id")), avcbrandid = Convert.ToString(p.Field<object>("avc品牌id")), kh = Convert.ToString(p.Field<object>("款号")) } into g
                                   select g;
                    if (queBrand.Count() > 0)
                    {
                        foreach (var q in queBrand)
                        {
                            MatchCollection mm2 = regex.Matches(q.Key.kh.ToUpper().Trim());
                            if (mm2.Count > 0)
                            {
                                haveabc = mm2[mm2.Count - 1].Value.ToString();
                                girardData2 = q.Key.kh.ToUpper().Trim().Substring(0, q.Key.kh.ToUpper().Trim().LastIndexOf(haveabc) + 1);
                            }
                            else
                                girardData2 = q.Key.kh.ToUpper().Trim();

                            //标准款号截取
                            GirardBrand = Regex.Split(girardData2, "~");

                            if (GirardData[0] == GirardBrand[0])
                            {
                                //开始计算 属性匹配度                                            
                                //属性相同=1，不同=0，NULL=空
                                //得分=0.25*吊牌价匹配+0.25*鞋帮高度匹配+0.25*闭合方式+0.25*性别
                                decimal codePrice = 0, codeHigh = 0, codeClose = 0, codeSex = 0;
                                if (dtTable.Rows[i]["鞋帮高度"].ToString().ToUpper().Trim() == q.Key.xbgd.ToUpper().Trim())
                                    codeHigh = 1;
                                if (dtTable.Rows[i]["闭合方式"].ToString().ToUpper().Trim() == q.Key.bhfs.ToUpper().Trim())
                                    codeClose = 1;
                                if (dtTable.Rows[i]["性别"].ToString().ToUpper().Trim() == q.Key.xb.ToUpper().Trim())
                                    codeSex = 1;
                                if (q.Key.dpj.ToUpper() != "NULL" && q.Key.dpj.ToUpper() != "" && dtTable.Rows[i]["吊牌价"].ToString().ToUpper().Trim() != "NULL" && dtTable.Rows[i]["吊牌价"].ToString().ToUpper().Trim() != "")
                                {
                                    if (Convert.ToDecimal(dtTable.Rows[i]["吊牌价"].ToString().ToUpper().Trim()) >= Convert.ToDecimal(q.Key.dpj))
                                    {
                                        if (Convert.ToDecimal(q.Key.dpj) * 100 / Convert.ToDecimal(dtTable.Rows[i]["吊牌价"].ToString().ToUpper().Trim()) >= 90)
                                            codePrice = 1;
                                    }
                                    else
                                    {
                                        if (Convert.ToDecimal(dtTable.Rows[i]["吊牌价"].ToString().ToUpper().Trim()) * 100 / Convert.ToDecimal(q.Key.dpj) >= 90)
                                            codePrice = 1;
                                    }
                                }
                                //最终匹配度
                                decimal endCode = Convert.ToDecimal(0.25) * (codePrice + codeHigh + codeClose + codeSex);
                                if (dtTable.Rows[i]["品类id"].ToString().ToUpper().Trim() == q.Key.plid.ToUpper().Trim())
                                {
                                    if (endCode >= Convert.ToDecimal(0.75))//匹配度>=0.75 则 判断（品类名称）是否相等
                                    {
                                        need = 0;
                                        dtBiaozhunData.Rows.Add(dtTable.Rows[i].ItemArray);
                                        dtBiaozhunData.Rows[dtBiaozhunData.Rows.Count - 1]["avc商品id"] = q.Key.avcplid;
                                        dtBiaozhunData.Rows[dtBiaozhunData.Rows.Count - 1]["avc品牌id"] = q.Key.avcbrandid;
                                        break;//跳出查找循环
                                    }
                                }
                                else
                                {
                                    if (endCode == Convert.ToDecimal(1))//匹配度>=0.75 则 判断（品类名称）是否相等
                                    {
                                        need = 0;
                                        dtBiaozhunData.Rows.Add(dtTable.Rows[i].ItemArray);
                                        dtBiaozhunData.Rows[dtBiaozhunData.Rows.Count - 1]["avc商品id"] = q.Key.avcplid;
                                        dtBiaozhunData.Rows[dtBiaozhunData.Rows.Count - 1]["avc品牌id"] = q.Key.avcbrandid;
                                        break;//跳出查找循环
                                    }
                                }
                            }
                        }
                    }
                }
                #endregion

                #endregion

                if (need == 1)
                {
                    var que = (from p in dtBiaozhunBrand.AsEnumerable() group p by Convert.ToString(p.Field<object>("avc商品id")) into g orderby Convert.ToInt32(g.Key.Replace("avc00_", "")) descending select g.Key).Take(1);
                    int add = 0;
                    foreach (var q in que)
                    {
                        add = Convert.ToInt32(q.Replace("avc00_", "")) + 1;
                    }
                    dtBiaozhunBrand.Rows.Add(dtTable.Rows[i].ItemArray);//添加为新标准品牌
                    dtBiaozhunBrand.Rows[dtBiaozhunBrand.Rows.Count - 1]["avc商品id"] = "avc00_" + add + "";
                    dtBiaozhunBrand.Rows[dtBiaozhunBrand.Rows.Count - 1]["avc品牌id"] = "avcb00_" + add + "";
                    //添加为标准数据
                    dtBiaozhunData.Rows.Add(dtTable.Rows[i].ItemArray);
                    dtBiaozhunData.Rows[dtBiaozhunData.Rows.Count - 1]["avc商品id"] = "avc00_" + add + "";
                    dtBiaozhunData.Rows[dtBiaozhunData.Rows.Count - 1]["avc品牌id"] = "avcb00_" + add + "";
                }
                #endregion
                this.Invoke(new ThreadStart(delegate()
                {
                    ProgressBarDisplay(dtTable, i);
                }));
            }
            //记录时间
            File.AppendAllText(@"timeText.txt", "【sheet--" + tableName + "】开始时间：" + strDatetime + " ----结束时间：" + DateTime.Now.ToString() + "\r\n");
            //导出数据
            //mySql.DataExportToFile(dtBiaozhunData, @"wgh" + tableName + ".xls");

            #endregion
        }

        private void TheForudMethod(DataTable dtTable, string tableName)
        {
            Debug.WriteLine("--" + DateTime.Now.ToString());
            for (int i = 0; i < dtTable.Rows.Count; i++)
            {
                var que = from p in dtTable.Select("品牌id='111111'").AsParallel() select p;
                Debug.WriteLine("end" + DateTime.Now.ToString());
                this.Invoke(new ThreadStart(delegate()
                {
                    ProgressBarDisplay(dtTable, i);
                }));
            }
            MessageBox.Show("w");
        }

        //女装
        private void NewMethodNV(DataTable dtTable, string tableName)
        {
            #region 测试（复杂方法）
            DataTable dtBiaozhunBrand = new DataTable();//存放标准品牌
            DataTable dtBiaozhunData = new DataTable();//存放查找到的标准数据
            //加载列名
            for (int i = 0; i < dtTable.Columns.Count; i++)
            {
                dtBiaozhunBrand.Columns.Add(dtTable.Columns[i].ColumnName);
                dtBiaozhunData.Columns.Add(dtTable.Columns[i].ColumnName);
            }

            //dtBiaozhunBrand、dtBiaozhunData在添加一个 标准字段，avc自定义（avc商品id、avc品牌id）
            dtBiaozhunBrand.Columns.Add("avc商品id", typeof(string));
            dtBiaozhunBrand.Columns.Add("avc型号id", typeof(string));
            dtBiaozhunBrand.Columns.Add("avc品牌id", typeof(string));
            dtBiaozhunData.Columns.Add("avc商品id", typeof(string));
            dtBiaozhunData.Columns.Add("avc型号id", typeof(string));
            dtBiaozhunBrand.Columns.Add("avc标准品牌", typeof(string));
            dtBiaozhunBrand.Columns.Add("avc品牌2", typeof(string));
            dtBiaozhunData.Columns.Add("avc标准品牌", typeof(string));
            dtBiaozhunData.Columns.Add("avc品牌2", typeof(string));
            dtBiaozhunData.Columns.Add("avc品牌id", typeof(string));

            for (int i = 0; i < dtTable.Rows.Count; i++)
            {
                Debug.WriteLine(i);

                if (i == 0)//第一条数据默认为标准数据
                {
                    // 拆分品牌
                    string[] Brand = dtTable.Rows[i]["品牌"].ToString().Split('/');
                    string strBrandChinase = string.Empty;
                    string strBrandEng = string.Empty;
                    for (int a = 0; a < Brand.Length; a++)
                    {
                        if (Brand.Length == 1)
                            strBrandChinase = strBrandEng = Brand[0];
                        else
                        {
                            //标准品牌都使用中文
                            Regex rgx = new Regex(@"[A-Za-z]");//全部为英文
                            if (rgx.IsMatch(Brand[a]))
                            {
                                strBrandEng = Brand[a];//英文
                            }
                            else
                                strBrandChinase = Brand[a];//中文品牌
                        }
                    }
                    dtBiaozhunBrand.Rows.Add(dtTable.Rows[i].ItemArray);//直接复制数据 行 
                    dtBiaozhunBrand.Rows[dtBiaozhunBrand.Rows.Count - 1]["avc标准品牌"] = strBrandChinase;
                    dtBiaozhunBrand.Rows[dtBiaozhunBrand.Rows.Count - 1]["avc品牌2"] = strBrandEng;
                    dtBiaozhunBrand.Rows[dtBiaozhunBrand.Rows.Count - 1]["avc商品id"] = "avc00_" + i + "";
                    dtBiaozhunBrand.Rows[dtBiaozhunBrand.Rows.Count - 1]["avc型号id"] = "Aowei00000" + (i + 1) + "";
                    dtBiaozhunBrand.Rows[dtBiaozhunBrand.Rows.Count - 1]["avc品牌id"] = "avcBrand00_" + i + "";
                    dtBiaozhunData.Rows.Add(dtTable.Rows[i].ItemArray);
                    dtBiaozhunData.Rows[dtBiaozhunData.Rows.Count - 1]["avc标准品牌"] = strBrandChinase;
                    dtBiaozhunData.Rows[dtBiaozhunData.Rows.Count - 1]["avc品牌2"] = strBrandEng;
                    dtBiaozhunData.Rows[dtBiaozhunData.Rows.Count - 1]["avc商品id"] = "avc00_" + i + "";
                    dtBiaozhunData.Rows[dtBiaozhunData.Rows.Count - 1]["avc型号id"] = "Aowei00000" + (i + 1) + "";
                    dtBiaozhunData.Rows[dtBiaozhunData.Rows.Count - 1]["avc品牌id"] = "avcBrand00_" + i + "";
                    continue;//继续下条数据
                }

                //开始逐条数据查找判断
                #region （从第二条数据开始） 开始逐条数据查找判断
                int need = 1;//查看该条数据是否在标准列表里面查找到.
                string strBrandChinase2 = string.Empty;
                string strBrandEng2 = string.Empty;
                for (int a = 0; a < dtBiaozhunBrand.Rows.Count; a++)
                {
                    // 拆分品牌
                    string[] Brand = dtTable.Rows[i]["品牌"].ToString().Split('/');

                    for (int ax = 0; ax < Brand.Length; ax++)
                    {
                        if (Brand.Length == 1)
                            strBrandChinase2 = strBrandEng2 = Brand[ax];
                        else
                        {
                            //标准品牌都使用中文
                            Regex rgx = new Regex(@"[A-Za-z]");//全部为英文
                            if (rgx.IsMatch(Brand[ax]))
                            {
                                strBrandEng2 = Brand[ax];//英文
                            }
                            else
                                strBrandChinase2 = Brand[ax];//中文品牌
                        }
                    }
                    #region //非特殊品牌（使用属性不同标记分割，吊牌价区别）
                    int codePrice = 0;
                    if (dtBiaozhunBrand.Rows[a]["品牌"].ToString().ToUpper().Trim().Contains(strBrandEng2.ToUpper().Trim()) || dtBiaozhunBrand.Rows[a]["品牌"].ToString().ToUpper().Trim().Contains(strBrandChinase2.ToUpper().Trim()))//品牌相等
                    {
                        if (dtTable.Rows[i]["衣长"].ToString().ToUpper().Trim() == dtBiaozhunBrand.Rows[a]["衣长"].ToString().ToUpper().Trim())//第2步 ：衣长 相等
                        {
                            if (dtTable.Rows[i]["袖长"].ToString().ToUpper().Trim() == dtBiaozhunBrand.Rows[a]["袖长"].ToString().ToUpper().Trim())//第3步 ：袖长 相等
                            {
                                if (dtTable.Rows[i]["图案"].ToString().ToUpper().Trim() == dtBiaozhunBrand.Rows[a]["图案"].ToString().ToUpper().Trim())//第4步 ：图案 相等
                                {
                                    if (dtTable.Rows[i]["标牌价"].ToString().ToUpper().Trim() == dtBiaozhunBrand.Rows[a]["标牌价"].ToString().ToUpper().Trim())
                                        codePrice = 1;
                                    if (dtTable.Rows[i]["标牌价"].ToString().ToUpper().Trim() != "NULL" && dtTable.Rows[i]["标牌价"].ToString().ToUpper().Trim() != "")
                                    {
                                        if (dtBiaozhunBrand.Rows[a]["标牌价"].ToString().ToUpper().Trim() != "NULL" && dtBiaozhunBrand.Rows[a]["标牌价"].ToString().ToUpper().Trim() != "")
                                        {
                                            if (Convert.ToDecimal(dtTable.Rows[i]["标牌价"].ToString().ToUpper().Trim()) > Convert.ToDecimal(dtBiaozhunBrand.Rows[a]["标牌价"].ToString().ToUpper().Trim()))
                                            {
                                                if (Convert.ToDecimal(dtBiaozhunBrand.Rows[a]["标牌价"].ToString().ToUpper().Trim()) * 100 / Convert.ToDecimal(dtTable.Rows[i]["标牌价"].ToString().ToUpper().Trim()) >= 90)
                                                    codePrice = 1;
                                            }
                                            else if (Convert.ToDecimal(dtTable.Rows[i]["标牌价"].ToString().ToUpper().Trim()) < Convert.ToDecimal(dtBiaozhunBrand.Rows[a]["标牌价"].ToString().ToUpper().Trim()))
                                            {
                                                if (Convert.ToDecimal(dtTable.Rows[i]["标牌价"].ToString().ToUpper().Trim()) * 100 / Convert.ToDecimal(dtBiaozhunBrand.Rows[a]["标牌价"].ToString().ToUpper().Trim()) >= 90)
                                                    codePrice = 1;
                                            }
                                        }
                                    }
                                    //标牌价相等
                                    if (codePrice == 1)
                                    {
                                        //  int endOK = 0;
                                        if (dtTable.Rows[i]["衣门襟"].ToString().ToUpper().Trim() == dtBiaozhunBrand.Rows[a]["衣门襟"].ToString().ToUpper().Trim())//第5步 ：衣门襟 相等
                                        {
                                            //判断属性
                                            //开始计算 属性匹配度                                            
                                            //属性相同=1，不同=0，NULL=空
                                            //得分= 货号+ 里料材质+ 厚薄+ 领子+风格+促销价
                                            decimal endOK = 0, codeHigh = 0, codeClose = 0, codeSex = 0, codeLengZi = 0, codeFengGe = 0;
                                            if (dtTable.Rows[i]["货号"].ToString().ToUpper().Trim() == dtBiaozhunBrand.Rows[a]["货号"].ToString().ToUpper().Trim())
                                                codeHigh = 1;
                                            if (dtTable.Rows[i]["里料材质"].ToString().ToUpper().Trim() == dtBiaozhunBrand.Rows[a]["里料材质"].ToString().ToUpper().Trim())
                                                codeClose = 1;
                                            if (dtTable.Rows[i]["厚薄"].ToString().ToUpper().Trim() == dtBiaozhunBrand.Rows[a]["厚薄"].ToString().ToUpper().Trim())
                                                codeSex = 1;
                                            if (dtTable.Rows[i]["领子"].ToString().ToUpper().Trim() == dtBiaozhunBrand.Rows[a]["领子"].ToString().ToUpper().Trim())
                                                codeLengZi = 1;
                                            if (dtTable.Rows[i]["风格"].ToString().ToUpper().Trim() == dtBiaozhunBrand.Rows[a]["风格"].ToString().ToUpper().Trim())
                                                codeFengGe = 1;
                                            if (dtTable.Rows[i]["促销价"].ToString().ToUpper().Trim() == dtBiaozhunBrand.Rows[a]["促销价"].ToString().ToUpper().Trim())
                                                endOK = 1;
                                            if (dtTable.Rows[i]["促销价"].ToString().ToUpper().Trim() != "NULL" && dtTable.Rows[i]["促销价"].ToString().ToUpper().Trim() != "")
                                            {
                                                if (dtBiaozhunBrand.Rows[a]["促销价"].ToString().ToUpper().Trim() != "NULL" && dtBiaozhunBrand.Rows[a]["促销价"].ToString().ToUpper().Trim() != "")
                                                {
                                                    if (Convert.ToDecimal(dtTable.Rows[i]["促销价"].ToString().ToUpper().Trim()) > Convert.ToDecimal(dtBiaozhunBrand.Rows[a]["促销价"].ToString().ToUpper().Trim()))
                                                    {
                                                        if (Convert.ToDecimal(dtBiaozhunBrand.Rows[a]["促销价"].ToString().ToUpper().Trim()) * 100 / Convert.ToDecimal(dtTable.Rows[i]["促销价"].ToString().ToUpper().Trim()) >= 90)
                                                            endOK = 1;
                                                    }
                                                    else if (Convert.ToDecimal(dtTable.Rows[i]["促销价"].ToString().ToUpper().Trim()) < Convert.ToDecimal(dtBiaozhunBrand.Rows[a]["促销价"].ToString().ToUpper().Trim()))
                                                    {
                                                        if (Convert.ToDecimal(dtTable.Rows[i]["促销价"].ToString().ToUpper().Trim()) * 100 / Convert.ToDecimal(dtBiaozhunBrand.Rows[a]["促销价"].ToString().ToUpper().Trim()) >= 90)
                                                            endOK = 1;
                                                    }
                                                }
                                            }
                                            //最终匹配度
                                            decimal endCode = Convert.ToDecimal(0.1667) * (endOK + codeHigh + codeClose + codeSex + codeLengZi + codeFengGe);//1/6
                                            if (endCode >= Convert.ToDecimal(0.8))//匹配度>=0.8 则 判断（品类名称）是否相等
                                            {
                                                need = 0;
                                                dtBiaozhunData.Rows.Add(dtTable.Rows[i].ItemArray);
                                                dtBiaozhunData.Rows[dtBiaozhunData.Rows.Count - 1]["avc标准品牌"] = strBrandChinase2;
                                                dtBiaozhunData.Rows[dtBiaozhunData.Rows.Count - 1]["avc品牌2"] = strBrandEng2;
                                                dtBiaozhunData.Rows[dtBiaozhunData.Rows.Count - 1]["avc商品id"] = dtBiaozhunBrand.Rows[a]["avc商品id"];
                                                dtBiaozhunData.Rows[dtBiaozhunData.Rows.Count - 1]["avc型号id"] = dtBiaozhunBrand.Rows[a]["avc型号id"];
                                                dtBiaozhunData.Rows[dtBiaozhunData.Rows.Count - 1]["avc品牌id"] = dtBiaozhunBrand.Rows[a]["avc品牌id"];
                                                break;//跳出查找循环
                                            }
                                            //else
                                            //{
                                            //    string brandaid = string.Empty;//avc型号id
                                            //    if ((i + 1) >= 10)
                                            //        brandaid = "Aowei0000" + (i+1) + "";
                                            //    else if ((i + 1) >= 100)
                                            //        brandaid = "Aowei000" + (i + 1) + "";
                                            //    else if ((i + 1) >= 1000)
                                            //        brandaid = "Aowei00" + (i + 1) + "";
                                            //    else if ((i + 1) >= 10000)
                                            //        brandaid = "Aowei0" + (i + 1) + "";

                                            //    dtBiaozhunBrand.Rows.Add(dtTable.Rows[i].ItemArray);//直接复制数据 行 
                                            //    dtBiaozhunBrand.Rows[dtBiaozhunBrand.Rows.Count - 1]["avc标准品牌"] = strBrandChinase2;
                                            //    dtBiaozhunBrand.Rows[dtBiaozhunBrand.Rows.Count - 1]["avc品牌2"] = strBrandEng2;
                                            //    dtBiaozhunBrand.Rows[dtBiaozhunBrand.Rows.Count - 1]["avc商品id"] = "avc00_" + i + "";
                                            //    dtBiaozhunBrand.Rows[dtBiaozhunBrand.Rows.Count - 1]["avc型号id"] = brandaid;
                                            //    dtBiaozhunBrand.Rows[dtBiaozhunBrand.Rows.Count - 1]["avc品牌id"] = "avcBrand00_" + i + "";
                                            //    dtBiaozhunData.Rows.Add(dtTable.Rows[i].ItemArray);
                                            //    dtBiaozhunData.Rows[dtBiaozhunData.Rows.Count - 1]["avc标准品牌"] = strBrandChinase2;
                                            //    dtBiaozhunData.Rows[dtBiaozhunData.Rows.Count - 1]["avc品牌2"] = strBrandEng2;
                                            //    dtBiaozhunData.Rows[dtBiaozhunData.Rows.Count - 1]["avc商品id"] = "avc00_" + i + "";
                                            //    dtBiaozhunData.Rows[dtBiaozhunData.Rows.Count - 1]["avc型号id"] = brandaid;
                                            //    dtBiaozhunData.Rows[dtBiaozhunData.Rows.Count - 1]["avc品牌id"] = "avcBrand00_" + i + "";
                                            //    break;
                                            //}
                                        }
                                    }
                                }
                            }
                        }
                    }
                    #endregion
                }
                if (need == 1)
                {
                    var que = (from p in dtBiaozhunBrand.AsEnumerable() group p by Convert.ToString(p.Field<object>("avc商品id")) into g orderby Convert.ToInt32(g.Key.Replace("avc00_", "")) descending select g.Key).Take(1);
                    int add = 0;
                    foreach (var q in que)
                    {
                        add = Convert.ToInt32(q.Replace("avc00_", "")) + 1;//商品id
                    }
                    //品牌id
                    var queBrandid = (from p in dtBiaozhunBrand.AsEnumerable() where Convert.ToString(p.Field<object>("avc标准品牌")).ToUpper() == strBrandChinase2.ToUpper() group p by Convert.ToString(p.Field<object>("avc品牌id")) into g orderby Convert.ToInt32(g.Key.Replace("avcBrand00_", "")) descending select g.Key).Take(1);
                    string addBrandid = string.Empty;
                    if (queBrandid.Count() > 0)
                    {
                        foreach (var q in queBrandid)
                            addBrandid = q;//品牌id
                    }
                    else
                    {
                        var queBrand = (from p in dtBiaozhunBrand.AsEnumerable() group p by Convert.ToString(p.Field<object>("avc品牌id")) into g orderby Convert.ToInt32(g.Key.Replace("avcBrand00_", "")) descending select g.Key).Take(1);
                        foreach (var q in queBrand)
                            addBrandid = "avcBrand00_" + (Convert.ToInt32(q.Replace("avcBrand00_", "")) + 1);//品牌id
                    }

                    //型号id
                    var queJxid = (from p in dtBiaozhunBrand.AsEnumerable() group p by Convert.ToString(p.Field<object>("avc型号id")) into g orderby Convert.ToInt32(g.Key.Replace("Aowei", "")) descending select g.Key).Take(1);
                    int addJxid = 0; string brandId = string.Empty;
                    foreach (var q in queJxid)
                        addJxid = Convert.ToInt32(q.Replace("Aowei", "")) + 1;//品牌id
                    if (addJxid < 10)
                        brandId = "Aowei00000" + addJxid + "";
                    else if (addJxid >= 10 && addJxid < 100)
                        brandId = "Aowei0000" + addJxid + "";
                    else if (addJxid >= 100 && addJxid < 1000)
                        brandId = "Aowei000" + addJxid + "";
                    else if (addJxid >= 1000 && addJxid < 10000)
                        brandId = "Aowei00" + addJxid + "";
                    else if (addJxid >= 10000)
                        brandId = "Aowei0" + addJxid + "";

                    dtBiaozhunBrand.Rows.Add(dtTable.Rows[i].ItemArray);//添加为新标准品牌
                    dtBiaozhunBrand.Rows[dtBiaozhunBrand.Rows.Count - 1]["avc标准品牌"] = strBrandChinase2;
                    dtBiaozhunBrand.Rows[dtBiaozhunBrand.Rows.Count - 1]["avc品牌2"] = strBrandEng2;
                    dtBiaozhunBrand.Rows[dtBiaozhunBrand.Rows.Count - 1]["avc商品id"] = "avc00_" + add + "";
                    dtBiaozhunBrand.Rows[dtBiaozhunBrand.Rows.Count - 1]["avc型号id"] = brandId;
                    dtBiaozhunBrand.Rows[dtBiaozhunBrand.Rows.Count - 1]["avc品牌id"] = addBrandid;
                    dtBiaozhunData.Rows.Add(dtTable.Rows[i].ItemArray);
                    dtBiaozhunData.Rows[dtBiaozhunData.Rows.Count - 1]["avc标准品牌"] = strBrandChinase2;
                    dtBiaozhunData.Rows[dtBiaozhunData.Rows.Count - 1]["avc品牌2"] = strBrandEng2;
                    dtBiaozhunData.Rows[dtBiaozhunData.Rows.Count - 1]["avc商品id"] = "avc00_" + add + "";
                    dtBiaozhunData.Rows[dtBiaozhunData.Rows.Count - 1]["avc型号id"] = brandId;
                    dtBiaozhunData.Rows[dtBiaozhunData.Rows.Count - 1]["avc品牌id"] = addBrandid;
                }
                #endregion
                this.Invoke(new ThreadStart(delegate()
                {
                    ProgressBarDisplay(dtTable, i);
                }));
            }
            //记录时间
            File.AppendAllText(@"timeText.txt", "【sheet--" + tableName + "】开始时间：" + strDatetime + " ----结束时间：" + DateTime.Now.ToString() + "\r\n");
            //导出数据

            path = AppDomain.CurrentDomain.BaseDirectory;//路径
            // mySql.DataExportToFile(dtBiaozhunData, @"avc" + tableName + ".xls");

            this.Invoke(new ThreadStart(delegate() { this.button2.Text = "打开:excel文件-" + AppDomain.CurrentDomain.BaseDirectory + "-"; }));
            #endregion
        }

        private void TddC(DataTable dtTable, string tableName)
        {
            DataTable dtTc = dtTable;
            dtTc = dtTc.AsEnumerable().Select(a =>
            {
                foreach (DataColumn dc in dtTc.Columns)
                {
                    if (a[dc].ToString() == "")
                        a[dc] = "NULL";
                }
                return a;
            }).CopyToDataTable();

            //dtTc = dtTc.AsEnumerable().Select(a =>
            //{

            //    foreach (DataColumn dc in dtTc.Columns)
            //    {
            //        a[dc] = a[dc].ToString().Replace("A", "B");//替换
            //    }
            //    return a;
            //}).CopyToDataTable<DataRow>();

            //记录时间
            File.AppendAllText(@"timeText.txt", "【sheet--" + tableName + "】开始时间：" + strDatetime + " ----结束时间：" + DateTime.Now.ToString() + "\r\n");
            //导出数据
            ////mySql.DataExportToFile(dtTc, @"套餐" + tableName + ".xls");//不在使用
            //
            //Debug.WriteLine(AppDomain.CurrentDomain.BaseDirectory+"套餐"+tableName+".xls");
            path = AppDomain.CurrentDomain.BaseDirectory;//路径
            mySql.DataOfGetExcel(dtTc, path + "套餐" + tableName + DateTime.Now.ToString().Replace("-", "").Replace(":", "").Replace("/", "") + ".xls");

            this.Invoke(new ThreadStart(delegate() { this.button2.Text = "打开:excel文件-" + AppDomain.CurrentDomain.BaseDirectory + "-"; }));

        }
        //线下整理套餐文件
        private void TC(DataTable dtTable, string tableName)
        {
            string[] Column = { "油烟机", "燃气灶", "热水器", "消毒柜" };
            DataTable dtTc = new DataTable();
            dtTc.Columns.Add("单品机型");
            dtTc.Columns.Add("单品品类");
            dtTc.Columns.Add("套餐机型");
            dtTc.Columns.Add("品类");
            dtTc.Columns.Add("品牌");
            for (int i = 0; i < dtTable.Rows.Count; i++)
            {
                //油烟机、燃气灶、热水器、消毒柜
                for (int j = 0; j < Column.Length; j++)
                {
                    DataRow dr = dtTc.NewRow();
                    dr["品类"] = "厨电套餐";
                    dr["单品品类"] = Column[j];
                    dr["单品机型"] = dtTable.Rows[i][Column[j]];
                    dr["套餐机型"] = dtTable.Rows[i]["机型"];
                    dr["品牌"] = dtTable.Rows[i]["品牌"];
                    if (dr["单品机型"].ToString() == "" || dr["单品机型"] == null)
                        continue;
                    dtTc.Rows.Add(dr);
                }
                this.Invoke(new ThreadStart(delegate()
                {
                    ProgressBarDisplay(dtTable, i);
                }));
            }
            //记录时间
            File.AppendAllText(@"timeText.txt", "【sheet--" + tableName + "】开始时间：" + strDatetime + " ----结束时间：" + DateTime.Now.ToString() + "\r\n");
            //导出数据
            ////mySql.DataExportToFile(dtTc, @"套餐" + tableName + ".xls");//不在使用
            //
            //Debug.WriteLine(AppDomain.CurrentDomain.BaseDirectory+"套餐"+tableName+".xls");
            path = AppDomain.CurrentDomain.BaseDirectory;//路径
            //Debug.WriteLine("be:" + DateTime.Now.ToString());
            //mySql.DataOfGetExcel(dtTc, path + "线下套餐" + tableName + DateTime.Now.ToString().Replace("-", "").Replace(":", "").Replace("/", "") + ".xls");
            //Debug.WriteLine("ed:" + DateTime.Now.ToString());
            DataOfGetExcel(dtTc, path + "线下套餐" + tableName + DateTime.Now.ToString().Replace("-", "").Replace(":", "").Replace("/", "") + ".xls");
            this.Invoke(new ThreadStart(delegate() { this.button2.Text = "打开:excel文件-" + AppDomain.CurrentDomain.BaseDirectory + "-"; }));

        }
        //线上套餐整理文件
        private void Online_TC(DataTable dtTable, string tableName)
        {
            string[] Column = { "油烟机", "燃气灶", "热水器", "消毒柜" };
            DataTable dtTc = new DataTable();
            dtTc.Columns.Add("单品机型");
            dtTc.Columns.Add("单品品类");
            dtTc.Columns.Add("套餐机型");
            dtTable.Columns.Add("单品品牌");
            dtTc.Columns.Add("品类");
            dtTc.Columns.Add("品牌");
            for (int i = 0; i < dtTable.Rows.Count; i++)
            {
                //油烟机、燃气灶、热水器、消毒柜
                for (int j = 0; j < Column.Length; j++)
                {
                    DataRow dr = dtTc.NewRow();
                    dr["品类"] = "厨电套餐";
                    dr["单品品类"] = Column[j];
                    dr["单品机型"] = dtTable.Rows[i][Column[j]];
                    dr["套餐机型"] = dtTable.Rows[i]["机型"];
                    dr["品牌"] = dtTable.Rows[i]["品牌"];
                    dr["单品品牌"] = dtTable.Rows[i]["单品品牌"];
                    if (dr["单品机型"].ToString() == "" || dr["单品机型"] == null)
                        continue;
                    dtTc.Rows.Add(dr);
                }
                this.Invoke(new ThreadStart(delegate()
                {
                    ProgressBarDisplay(dtTable, i);
                }));
            }
            //记录时间
            File.AppendAllText(@"timeText.txt", "【sheet--" + tableName + "】开始时间：" + strDatetime + " ----结束时间：" + DateTime.Now.ToString() + "\r\n");
            path = AppDomain.CurrentDomain.BaseDirectory;//路径
            DataOfGetExcel(dtTc, path + "线上套餐" + tableName + DateTime.Now.ToString().Replace("-", "").Replace(":", "").Replace("/", "") + ".xls");

            this.Invoke(new ThreadStart(delegate() { this.button2.Text = "打开:excel文件-" + AppDomain.CurrentDomain.BaseDirectory + "-"; }));

        }

        //整理冰箱数据
        private void TheFRMethod(DataTable dtTable, string tableName)
        {
            //商品id、商品名称、店铺名称、品牌id、品牌名称、品类id、品类名称、吊牌价、运动鞋分类 
            #region 测试（复杂方法）
            DataTable dtBiaozhunBrand = new DataTable();//存放标准品牌
            DataTable dtBiaozhunData = new DataTable();//存放查找到的标准数据
            //加载列名
            for (int i = 0; i < dtTable.Columns.Count; i++)
            {
                dtBiaozhunBrand.Columns.Add(dtTable.Columns[i].ColumnName);
                dtBiaozhunData.Columns.Add(dtTable.Columns[i].ColumnName);
            }
            dtBiaozhunBrand.Columns.Add("avc商品id", typeof(string));
            dtBiaozhunData.Columns.Add("avc商品id", typeof(string));

            for (int i = 0; i < dtTable.Rows.Count; i++)
            {
                if (i == 0)//第一条数据默认为标准数据
                {
                    dtBiaozhunBrand.Rows.Add(dtTable.Rows[i].ItemArray);//直接复制数据 行 
                    dtBiaozhunBrand.Rows[dtBiaozhunBrand.Rows.Count - 1]["avc商品id"] = "" + i + "";
                    dtBiaozhunData.Rows.Add(dtTable.Rows[i].ItemArray);
                    dtBiaozhunData.Rows[dtBiaozhunData.Rows.Count - 1]["avc商品id"] = "" + i + "";
                    continue;//继续下条数据
                }

                //开始逐条数据查找判断
                #region （从第二条数据开始） 开始逐条数据查找判断
                int need = 1;//查看该条数据是否在标准列表里面查找到.
                for (int a = 0; a < dtBiaozhunBrand.Rows.Count; a++)
                {
                    decimal codeSize = 0, codeWeek = 0;
                    #region //非特殊品牌（使用属性不同标记分割，吊牌价区别）

                    if (dtTable.Rows[i]["品牌"].ToString().ToUpper().Trim() == dtBiaozhunBrand.Rows[a]["品牌"].ToString().ToUpper().Trim())//1、品牌
                    {
                        if (dtTable.Rows[i]["门数"].ToString().ToUpper().Trim() == dtBiaozhunBrand.Rows[a]["门数"].ToString().ToUpper().Trim())//2、门数
                        {
                            if (dtTable.Rows[i]["温控方式"].ToString().ToUpper().Trim() == dtBiaozhunBrand.Rows[a]["温控方式"].ToString().ToUpper().Trim())//3、温控方式
                            {
                                if (dtTable.Rows[i]["定变频"].ToString().ToUpper().Trim() == dtBiaozhunBrand.Rows[a]["定变频"].ToString().ToUpper().Trim())//4、定变频
                                {
                                    if (dtTable.Rows[i]["玻璃面板"].ToString().ToUpper().Trim() == dtBiaozhunBrand.Rows[a]["玻璃面板"].ToString().ToUpper().Trim())//5、玻璃面板
                                    {
                                        if (dtTable.Rows[i]["制冷方式"].ToString().ToUpper().Trim() == dtBiaozhunBrand.Rows[a]["制冷方式"].ToString().ToUpper().Trim())//6、制冷方式
                                        {
                                            if (dtTable.Rows[i]["容积"].ToString().ToUpper().Trim() != "NULL" && dtTable.Rows[i]["容积"].ToString().ToUpper().Trim() != "")//7、比容积，5%的容差
                                            {
                                                if (dtBiaozhunBrand.Rows[a]["容积"].ToString().ToUpper().Trim() != "NULL" && dtBiaozhunBrand.Rows[a]["容积"].ToString().ToUpper().Trim() != "")
                                                {
                                                    if (Convert.ToDecimal(dtTable.Rows[i]["容积"].ToString().ToUpper().Trim()) >= Convert.ToDecimal(dtBiaozhunBrand.Rows[a]["容积"].ToString().ToUpper().Trim()))
                                                    {
                                                        if (Convert.ToDecimal(dtBiaozhunBrand.Rows[a]["容积"].ToString().ToUpper().Trim()) * 100 / Convert.ToDecimal(dtTable.Rows[i]["容积"].ToString().ToUpper().Trim()) >= 95)
                                                            codeSize = 1;
                                                    }
                                                    else if (Convert.ToDecimal(dtTable.Rows[i]["容积"].ToString().ToUpper().Trim()) < Convert.ToDecimal(dtBiaozhunBrand.Rows[a]["容积"].ToString().ToUpper().Trim()))
                                                    {
                                                        if (Convert.ToDecimal(dtTable.Rows[i]["容积"].ToString().ToUpper().Trim()) * 100 / Convert.ToDecimal(dtBiaozhunBrand.Rows[a]["容积"].ToString().ToUpper().Trim()) >= 95)
                                                            codeSize = 1;
                                                    }
                                                }
                                            }
                                            if (dtTable.Rows[i]["上市周度"].ToString().ToUpper().Trim() != "NULL" && dtTable.Rows[i]["上市周度"].ToString().ToUpper().Trim() != "" && dtTable.Rows[i]["上市周度"].ToString().ToUpper().Trim() != "-")//8、对比上市周度，6个月的容差 4*6=24周
                                            {
                                                if (dtBiaozhunBrand.Rows[a]["上市周度"].ToString().ToUpper().Trim() != "NULL" && dtBiaozhunBrand.Rows[a]["上市周度"].ToString().ToUpper().Trim() != "" && dtBiaozhunBrand.Rows[a]["上市周度"].ToString().ToUpper().Trim() != "-")
                                                {
                                                    if (Math.Abs(Convert.ToDecimal(dtBiaozhunBrand.Rows[a]["上市周度"].ToString().ToUpper().Trim().Replace("W", "")) - Convert.ToDecimal(dtTable.Rows[i]["上市周度"].ToString().ToUpper().Trim().Replace("W", ""))) <= 25)//8、对比上市周度，6个月的容差 4*6=24周
                                                        codeWeek = 1;
                                                }
                                            }
                                            else if (dtTable.Rows[i]["上市周度"].ToString().ToUpper().Trim() == "-")
                                            {
                                                if (dtTable.Rows[i]["上市周度"].ToString().ToUpper().Trim() == dtBiaozhunBrand.Rows[a]["上市周度"].ToString().ToUpper().Trim())
                                                    codeWeek = 1;
                                            }

                                            if (codeSize == 1 && codeWeek == 1)
                                            {
                                                need = 0;
                                                dtBiaozhunData.Rows.Add(dtTable.Rows[i].ItemArray);
                                                dtBiaozhunData.Rows[dtBiaozhunData.Rows.Count - 1]["avc商品id"] = dtBiaozhunBrand.Rows[a]["avc商品id"];
                                                break;//跳出查找循环
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                    #endregion
                }
                if (need == 1)
                {
                    var que = (from p in dtBiaozhunBrand.AsEnumerable() group p by Convert.ToString(p.Field<object>("avc商品id")) into g orderby Convert.ToInt32(g.Key.Replace("avc00_", "")) descending select g.Key).Take(1);
                    int add = 0;
                    foreach (var q in que)
                    {
                        add = Convert.ToInt32(q.Replace("avc00_", "")) + 1;
                    }
                    dtBiaozhunBrand.Rows.Add(dtTable.Rows[i].ItemArray);//添加为新标准品牌
                    dtBiaozhunBrand.Rows[dtBiaozhunBrand.Rows.Count - 1]["avc商品id"] = "" + add + "";
                    //添加为标准数据
                    dtBiaozhunData.Rows.Add(dtTable.Rows[i].ItemArray);
                    dtBiaozhunData.Rows[dtBiaozhunData.Rows.Count - 1]["avc商品id"] = "" + add + "";
                }
                #endregion
                this.Invoke(new ThreadStart(delegate()
                {
                    ProgressBarDisplay(dtTable, i);
                }));
            }
            //记录时间
            File.AppendAllText(@"timeText.txt", "【sheet--" + tableName + "】开始时间：" + strDatetime + " ----结束时间：" + DateTime.Now.ToString() + "\r\n");
            //导出数据
            try
            {
                // mySql.DataExportToFile(dtBiaozhunData, @"奥维数据分析-" + tableName + ".xls");
            }
            catch (Exception e)
            {
                label1.Invoke(new ThreadStart(delegate() { label1.Text += "" + e.Message + ""; }));
            }
            finally
            {
                label1.Invoke(new ThreadStart(delegate() { label1.Text += "导出完成！"; }));
            }
            #endregion
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //this.Close();
            Application.Exit();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (path.Length > 0)
                Process.Start(path);
            else
            {
                MessageBox.Show("888888888888888888888888888", "it"); return;
            }
        }

        /// <summary>
        /// 写入excel
        /// </summary>
        /// <param name="dtTc"></param>
        /// <param name="path"></param>
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
            this.progressBar1.Invoke(new ThreadStart(delegate() { this.progressBar1.Minimum = 0; this.progressBar1.Maximum = dtTc.Rows.Count; }));
            this.label1.Invoke(new ThreadStart(delegate() { this.label1.Text = "开始写入Excel文件："; }));
            int r = 0, c = 0;
            for (int i = 0; i < dtTc.Rows.Count; i++)
            {
                this.progressBar1.Invoke(new ThreadStart(delegate() { this.progressBar1.Value = (i + 1); }));
                this.label1.Invoke(new ThreadStart(delegate() { this.label1.Text = "开始写入Excel文件：开始  " + (i + 1) + " 共 " + dtTc.Rows.Count + ""; }));
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

        //处理省份、地市、县市、区 4级联
        private void openExcelSheets(DataTable dtx, string filePath)
        {
            dtResultSheetCount = new DataTable();
            dtResultSheetCount.Columns.Add("省份");
            dtResultSheetCount.Columns.Add("地市");
            dtResultSheetCount.Columns.Add("县市");
            dtResultSheetCount.Columns.Add("街道");
            dtResultSheetCount.Columns.Add("省份id"); ;
            dtResultSheetCount.Columns.Add("地市id");
            dtResultSheetCount.Columns.Add("县市id");
            dtResultSheetCount.Columns.Add("街道id");
            dtResultSheetCount.Columns.Add("省份缩写");
            dtResultSheetCount.Columns.Add("地市缩写");
            dtResultSheetCount.Columns.Add("县市缩写");
            dtResultSheetCount.Columns.Add("街道缩写");
            //读取excel
            //先遍历sheet个数
            List<string> workSheet = new List<string>();
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            excelApp.DisplayAlerts = false;
            excelApp.Visible = false;
            Microsoft.Office.Interop.Excel.Workbook book = excelApp.Workbooks.Open(filePath);
            foreach (Microsoft.Office.Interop.Excel.Worksheet sheet in book.Worksheets)
                workSheet.Add(sheet.Name);
            //Debug.WriteLine(filePath + "--" + workSheet.Count);
            //1、先查询省份
            var queprovince = (from p in dtx.Select("parentid=0").AsParallel()
                               group p by new { id = Convert.ToString(p.Field<object>("id")), province = Convert.ToString(p.Field<object>("areaname")), suox = p.Field<string>("shortname") } into g
                               orderby g.Key.province
                               select g);
            foreach (var q in queprovince)
            {
                string sf = string.Empty, sfsx = string.Empty, sfid = string.Empty, ds = string.Empty, dssx = string.Empty, dsid = string.Empty, xs = string.Empty, xssx = string.Empty, xsid = string.Empty;
                sf = q.Key.province;
                sfid = q.Key.id;
                sfsx = q.Key.suox;
                //查找地市
                var qcity = (from p in dtx.Select("parentid=" + q.Key.id + "").AsParallel()
                             group p by new { id = p.Field<object>("id"), city = p.Field<string>("areaname"), suox = p.Field<string>("shortname") } into g
                             select g.Key);
                if (qcity.Count() <= 0)
                {
                    DataRow dr = dtResultSheetCount.NewRow();
                    dr["省份"] = sf;
                    dr["省份缩写"] = sfsx;
                    dr["省份id"] = sfid;
                    dr["地市"] = ds;
                    dr["地市缩写"] = dssx;
                    dr["地市id"] = dsid;
                    dr["县市"] = xs;
                    dr["县市缩写"] = xssx;
                    dr["县市id"] = xsid;
                    dtResultSheetCount.Rows.Add(dr);
                    continue;
                }

                foreach (var qc in qcity)
                {
                    ds = qc.city;
                    dsid = qc.id.ToString();
                    dssx = qc.suox;
                    //查找县市
                    var qxianshi = (from p in dtx.Select("parentid=" + qc.id + "").AsParallel()
                                    group p by new { id = p.Field<object>("id"), city = p.Field<string>("areaname"), suox = p.Field<string>("shortname") } into g
                                    select g.Key);
                    foreach (var qx in qxianshi)
                    {
                        xs = qx.city;
                        xsid = qx.id.ToString();
                        xssx = qx.suox;
                        //查询街道
                        var qjiedao = (from p in dtx.Select("parentid=" + qx.id + "").AsParallel()
                                       group p by new { id = p.Field<object>("id"), city = p.Field<string>("areaname"), suox = p.Field<string>("shortname") } into g
                                       select g.Key);
                        if (qjiedao.Count() <= 0)
                        {
                            DataRow dr = dtResultSheetCount.NewRow();
                            dr["省份"] = sf;
                            dr["省份缩写"] = sfsx;
                            dr["省份id"] = sfid;
                            dr["地市"] = ds;
                            dr["地市缩写"] = dssx;
                            dr["地市id"] = dsid;
                            dr["县市"] = xs;
                            dr["县市缩写"] = xssx;
                            dr["县市id"] = xsid;
                            dtResultSheetCount.Rows.Add(dr);
                            continue;
                        }
                        foreach (var qj in qjiedao)
                        {
                            DataRow dr = dtResultSheetCount.NewRow();
                            dr["省份"] = sf;
                            dr["省份缩写"] = sfsx;
                            dr["省份id"] = sfid;
                            dr["地市"] = ds;
                            dr["地市缩写"] = dssx;
                            dr["地市id"] = dsid;
                            dr["县市"] = xs;
                            dr["县市缩写"] = xssx;
                            dr["县市id"] = xsid;
                            dr["街道"] = qj.city;
                            dr["街道缩写"] = qj.suox;
                            dr["街道id"] = qj.id;
                            var qxxx = (from p in dtx.Select("parentid=" + qj.id + "").AsParallel()
                                        group p by new { id = p.Field<object>("id"), city = p.Field<string>("areaname"), suox = p.Field<string>("shortname") } into g
                                        select g.Key);
                            dtResultSheetCount.Rows.Add(dr);
                            if (qxxx.Count() > 0)
                            {
                                foreach (var qxj in qxxx)
                                {
                                    Debug.WriteLine(qxj.city);
                                }
                            }
                        }
                    }
                }

            }
            //for (int i = 0; i < dtx.Rows.Count; i++)
            //{
            //    DataRow dr = dtResultSheetCount.NewRow();

            //}
            //关闭excel进程           
            //try
            //{
            //    book.Close();
            //    excelApp.Quit();
            //    if (excelApp != null)
            //    {
            //        //获取Excel App的句柄
            //        hwnd = new IntPtr(excelApp.Hwnd);
            //        //通过Windows API获取Excel进程ID
            //        GetWindowThreadProcessId(hwnd, out pid);
            //        if (pid > 0)
            //        {
            //            Process process = Process.GetProcessById(pid);
            //            process.Kill();
            //        }
            //    }
            //}
            //catch (Exception e)
            //{
            //    MessageBox.Show(e.Message);
            //}
        }

        /// <summary>
        /// 查找线下周度最低价，2015-2016年
        /// </summary>
        /// <param name="dtx"></param>
        /// <param name="filePath"></param>
        private void SelectYJBLowerPrice_old(DataTable dtx, string filePath)
        {
            dtResultSheetCount = new DataTable();
            dtResultSheetCount.Columns.Add("SKU名称");
            dtResultSheetCount.Columns.Add("品牌");
            dtResultSheetCount.Columns.Add("历史最低售价");
            dtResultSheetCount.Columns.Add("平均售价");
            dtResultSheetCount.Columns.Add("品类");
            dtResultSheetCount.Columns.Add("机型");
            //读取excel
            //先遍历sheet个数
            List<string> workSheet = new List<string>();
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            excelApp.DisplayAlerts = false;
            excelApp.Visible = false;
            Microsoft.Office.Interop.Excel.Workbook book = excelApp.Workbooks.Open(filePath);
            foreach (Microsoft.Office.Interop.Excel.Worksheet sheet in book.Worksheets)
                workSheet.Add(sheet.Name);
            //查找数据库的 型号表
            string sxl = "SELECT a.品类,品牌,机型,组别简称 FROM 型号表 A JOIN 品类表 B ON a.品类=B.品类";
            DataTable dtJx = mySql.GetdtTable(sxl);

            for (int i = 0; i < dtx.Rows.Count; i++)
            {
                this.label1.Invoke(new ThreadStart(delegate() { this.label1.Text = "当前进度：共 " + dtx.Rows.Count + " 条 / 开始 " + (i + 1) + " "; }));

                var queJx = (from p in dtJx.AsEnumerable()
                             where dtx.Rows[i]["SKU名称"].ToString().ToUpper().Contains(p.Field<string>("机型").ToUpper()) && dtx.Rows[i]["品牌"].ToString().ToUpper() == p.Field<string>("品牌").ToUpper()
                             select new { jx = p.Field<string>("机型"), pl = p.Field<string>("品类"), pljc = p.Field<string>("组别简称") });
                if (queJx.Count() >= 2)
                {
                    //foreach (var q in queJx)
                    //{
                    //    Debug.WriteLine(q.jx);
                    //}
                }
                if (queJx.Count() > 0)
                {
                    string jx = string.Empty, pl = string.Empty, pljc = string.Empty;
                    if (queJx.Count() >= 2)
                    {
                        foreach (var q in queJx)
                        {
                            if (q.jx.Length > jx.Length)
                                jx = q.jx;
                            pl = q.pl;
                            pljc = q.pljc;
                        }
                    }
                    else
                    {
                        foreach (var q in queJx)
                        {
                            jx = q.jx;
                            pl = q.pl;
                            pljc = q.pljc;
                        }
                    }

                    DataRow dr = dtResultSheetCount.NewRow();
                    string yjb = "" +pljc + "_线下周度" + pl + "永久表2016年", lstyjb = "" + pljc + "_线下周度" + pl + "永久表2015年";
                    sxl = "SELECT MIN(单价)单价 FROM (" +
                        "SELECT 单价 FROM " + yjb + " WHERE 品牌='" + dtx.Rows[i]["品牌"] + "'  AND 机型='" + jx + "' " +
                        "UNION  " +
                        " SELECT 单价 FROM " + lstyjb + " WHERE 品牌='" + dtx.Rows[i]["品牌"] + "'  AND 机型='" +jx + "')A ";
                    DataTable dtlowPrice = mySql.GetdtTable(sxl);
                    //均价
                    sxl = "SELECT SUM(销额)/SUM(销量)  FROM (" +
                        "SELECT SUM(销量)销量,SUM(销额)销额 FROM " + yjb + " WHERE 品牌='" + dtx.Rows[i]["品牌"] + "'  AND 机型='" + jx + "' " +
                        "UNION  " +
                        " SELECT SUM(销量)销量,SUM(销额)销额 FROM " + lstyjb + " WHERE 品牌='" + dtx.Rows[i]["品牌"] + "'  AND 机型='" + jx + "')A ";
                    DataTable dtPrice = mySql.GetdtTable(sxl);
                    if (dtlowPrice.Rows.Count >= 0)
                    {
                        if (dtlowPrice.Rows[0][0] != DBNull.Value)
                            dr["历史最低售价"] = dtlowPrice.Rows[0][0];
                        if (dtPrice.Rows[0][0] != DBNull.Value)
                            dr["平均售价"] = dtPrice.Rows[0][0];
                    }
                    else
                    {
                        dr["历史最低售价"] = 0;
                        dr["平均售价"] = 0;
                    }
                    dr["SKU名称"] = dtx.Rows[i]["SKU名称"];
                    dr["品牌"] = dtx.Rows[i]["品牌"];
                    dr["品类"] = pl;
                    dr["机型"] = jx;
                    dtResultSheetCount.Rows.Add(dr);
                }
            }
        }

        private void SelectYJBLowerPrice(DataTable dtx, string filePath)
        {
            dtResultSheetCount = new DataTable();
            dtResultSheetCount.Columns.Add("SKU名称");
            dtResultSheetCount.Columns.Add("品牌");
            dtResultSheetCount.Columns.Add("历史最低售价");
            dtResultSheetCount.Columns.Add("平均售价");
            dtResultSheetCount.Columns.Add("品类");
            dtResultSheetCount.Columns.Add("机型");
            dtResultConnig = dtResultSheetCount.Clone();//临时表
            //读取excel
            //先遍历sheet个数
            List<string> workSheet = new List<string>();
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            excelApp.DisplayAlerts = false;
            excelApp.Visible = false;
            Microsoft.Office.Interop.Excel.Workbook book = excelApp.Workbooks.Open(filePath);
            foreach (Microsoft.Office.Interop.Excel.Worksheet sheet in book.Worksheets)
                workSheet.Add(sheet.Name);
            //查找数据库的 型号表
            string sxl = "SELECT a.品类,品牌,机型,组别简称 FROM 型号表 A JOIN 品类表 B ON a.品类=B.品类";
            DataTable dtJx = mySql.GetdtTable(sxl);

            for (int i = 0; i < dtx.Rows.Count; i++)
            {
                if (dtx.Rows[i]["SKU名称"].ToString().ToUpper().Contains("LJSQ20-12U1(世纪星)*12T RQ12U1DB2(C)(万家乐)"))
                    Debug.WriteLine(DateTime.Now.ToString());
                this.label1.Invoke(new ThreadStart(delegate() { this.label1.Text = "当前进度：共 " + dtx.Rows.Count + " 条 / 开始 " + (i + 1) + " "; }));
                string spmc = dtx.Rows[i]["SKU名称"].ToString().ToUpper(), slx = string.Empty;
                if (dtx.Rows[i]["SKU名称"].ToString().ToUpper().StartsWith("JZT-") || dtx.Rows[i]["SKU名称"].ToString().ToUpper().StartsWith("JZY-"))
                {
                    spmc = spmc.Replace("JZT-", "JZ(Y.R.T)2-").Replace("JZT-", "JZ(Y.T.R)2-").Replace("JZY-", "JZ(Y.R.T)2-").ToUpper();
                    slx = "品类='燃气灶'";
                }
                var queJx = (from p in dtJx.Select(slx).AsParallel()
                             where spmc.Contains(p.Field<string>("机型").ToUpper()) && dtx.Rows[i]["品牌"].ToString().ToUpper() == p.Field<string>("品牌").ToUpper()
                             select new { jx = p.Field<string>("机型"), pl = p.Field<string>("品类"), pljc = p.Field<string>("组别简称") });
                #region 处理
                if (queJx.Count() > 0)
                {
                    string jx = string.Empty, pl = string.Empty, pljc = string.Empty;
                    if (queJx.Count() >= 2)
                    {
                        foreach (var q in queJx)
                        {
                            if (q.jx.Length > jx.Length)
                                jx = q.jx;
                            pl = q.pl;
                            pljc = q.pljc;
                        }
                    }
                    else
                    {
                        foreach (var q in queJx)
                        {
                            jx = q.jx;
                            pl = q.pl;
                            pljc = q.pljc;
                        }
                    }
                    //foreach (var q in queJx)
                    //{
                    DataRow dr = dtResultSheetCount.NewRow();
                    dr["历史最低售价"] = 0;
                    dr["平均售价"] = 0;
                    dr["SKU名称"] = dtx.Rows[i]["SKU名称"];
                    dr["品牌"] = dtx.Rows[i]["品牌"];
                    dr["品类"] = pl;
                    dr["机型"] = jx;
                    dtResultSheetCount.Rows.Add(dr);
                    //}
                }
                else
                {
                    DataRow dr = dtResultSheetCount.NewRow();
                    dr["历史最低售价"] = 0;
                    dr["SKU名称"] = dtx.Rows[i]["SKU名称"];
                    dr["品牌"] = dtx.Rows[i]["品牌"];
                    dr["平均售价"] = 0;
                    dr["品类"] = "-";
                    dr["机型"] = "-";
                    dtResultSheetCount.Rows.Add(dr);
                }
                #endregion
            }
            #region 新处理

            IEnumerable<DataRow> quex = from p in dtResultSheetCount.AsEnumerable() orderby p.Field<string>("品类") descending select p;
            if (quex.Count() > 0)
            {
                dtResultConnig = quex.CopyToDataTable();
                dtResultSheetCount.Clear();
            }

            string xplsx = string.Empty;
            var plsx = (from p in dtJx.Select("品类='" + dtResultConnig.Rows[0]["品类"] + "'").AsParallel()
                        select p.Field<string>("组别简称")).Take(1);
            foreach (var q in plsx)
                xplsx = q.ToString();

            string yjb = "" + xplsx + "_线下周度" + dtResultConnig.Rows[0]["品类"].ToString() + "永久表2016年", lstyjb = "" + xplsx + "_线下周度" + dtResultConnig.Rows[0]["品类"].ToString() + "永久表2015年";
            sxl = "SELECT MIN(单价)单价,品牌,机型 FROM (" +
                "SELECT MIN(单价)单价,品牌,机型 FROM " + yjb + " GROUP BY 机型,品牌 " +
                "UNION  " +
                " SELECT MIN(单价)单价,品牌,机型 FROM " + lstyjb + " GROUP BY 机型,品牌)A GROUP By 品牌,机型";
            DataTable dtlowPrice = mySql.GetdtTable(sxl);
            //均价
            string sxl2 = "SELECT SUM(销额)/SUM(销量)单价,品牌,机型  FROM (" +
                "SELECT SUM(销量)销量,SUM(销额)销额,品牌,机型 FROM " + yjb + " GROUP BY 品牌,机型 " +
                "UNION  " +
                " SELECT SUM(销量)销量,SUM(销额)销额,品牌,机型 FROM " + lstyjb + " group by 品牌,机型)A group by 品牌,机型 ";
            DataTable dtPrice = mySql.GetdtTable(sxl2);
            for (int i = 0; i < dtResultConnig.Rows.Count; i++)
            {
                this.label1.Invoke(new ThreadStart(delegate() { this.label1.Text = "开始处理：" + dtResultConnig.Rows.Count + " / " + (i + 1) + ""; }));
                if (dtResultConnig.Rows[i]["品类"].ToString().ToUpper() == "-")
                {
                    DataRow dr = dtResultSheetCount.NewRow();
                    dr["历史最低售价"] = 0;
                    dr["平均售价"] = 0;
                    dr["SKU名称"] = dtResultConnig.Rows[i]["SKU名称"];
                    dr["品牌"] = dtResultConnig.Rows[i]["品牌"];
                    dr["品类"] = "-";
                    dr["机型"] = "-";
                    dtResultSheetCount.Rows.Add(dr);
                    continue;
                }
                if (i < dtResultConnig.Rows.Count - 1 && i > 0)
                {
                    if (dtResultConnig.Rows[i]["品类"].ToString().ToUpper() == dtResultConnig.Rows[i - 1]["品类"].ToString().ToUpper() && dtResultConnig.Rows[i]["品类"].ToString().ToUpper() != "-")
                    {

                    }
                    else
                    {
                        var plsxx = (from p in dtJx.Select("品类='" + dtResultConnig.Rows[i]["品类"] + "'").AsParallel()
                                     select p.Field<string>("组别简称")).Take(1);
                        foreach (var q in plsxx)
                            xplsx = q.ToString();
                        yjb = "" + xplsx + "_线下周度" + dtResultConnig.Rows[i]["品类"].ToString() + "永久表2016年";
                        lstyjb = "" + xplsx + "_线下周度" + dtResultConnig.Rows[i]["品类"].ToString() + "永久表2015年";
                        sxl = "SELECT MIN(单价)单价,品牌,机型 FROM (" +
                            "SELECT MIN(单价)单价,品牌,机型 FROM " + yjb + " GROUP BY 机型,品牌 " +
                            "UNION  " +
                            " SELECT MIN(单价)单价,品牌,机型 FROM " + lstyjb + " GROUP BY 机型,品牌)A GROUP By 品牌,机型";
                        dtlowPrice = mySql.GetdtTable(sxl);
                        sxl2 = "SELECT SUM(销额)/SUM(销量)单价,品牌,机型  FROM (" +
                    "SELECT SUM(销量)销量,SUM(销额)销额,品牌,机型 FROM " + yjb + " GROUP BY 品牌,机型 " +
                    "UNION  " +
                    " SELECT SUM(销量)销量,SUM(销额)销额,品牌,机型 FROM " + lstyjb + " group by 品牌,机型)A group by 品牌,机型 ";
                        dtPrice = mySql.GetdtTable(sxl2);
                    }
                    DataRow dr = dtResultSheetCount.NewRow();
                    var queJx = (from p in dtlowPrice.Select("品牌='" + dtResultConnig.Rows[i]["品牌"] + "' AND 机型='" + dtResultConnig.Rows[i]["机型"] + "'")
                                 select Convert.ToDecimal(p.Field<object>("单价")));
                    if (queJx.Count() > 0)
                    {
                        foreach (var q in queJx)
                        {
                            dr["历史最低售价"] = q;
                        }
                    }
                    else
                        dr["历史最低售价"] = 0;
                    var queJp = (from p in dtPrice.Select("品牌='" + dtResultConnig.Rows[i]["品牌"] + "' AND 机型='" + dtResultConnig.Rows[i]["机型"] + "'")
                                 select Convert.ToDecimal(p.Field<object>("单价")));
                    if (queJp.Count() > 0)
                    {
                        foreach (var q in queJp)
                            dr["平均售价"] = q;
                    }
                    else
                        dr["平均售价"] = 0;

                    dr["SKU名称"] = dtResultConnig.Rows[i]["SKU名称"];
                    dr["品牌"] = dtResultConnig.Rows[i]["品牌"];
                    dr["品类"] = dtResultConnig.Rows[i]["品类"].ToString();
                    dr["机型"] = dtResultConnig.Rows[i]["机型"];
                    dtResultSheetCount.Rows.Add(dr);
                }
                else if (i == 0)
                {
                    DataRow dr = dtResultSheetCount.NewRow();
                    var queJx = (from p in dtlowPrice.Select("品牌='" + dtResultConnig.Rows[i]["品牌"] + "' AND 机型='" + dtResultConnig.Rows[i]["机型"] + "'")
                                 select Convert.ToDecimal(p.Field<object>("单价")));
                    if (queJx.Count() > 0)
                    {
                        foreach (var q in queJx)
                        {
                            dr["历史最低售价"] = q;
                        }
                    }
                    else
                        dr["历史最低售价"] = 0;
                    var queJp = (from p in dtPrice.Select("品牌='" + dtResultConnig.Rows[i]["品牌"] + "' AND 机型='" + dtResultConnig.Rows[i]["机型"] + "'")
                                 select Convert.ToDecimal(p.Field<object>("单价")));
                    if (queJp.Count() > 0)
                    {
                        foreach (var q in queJp)
                            dr["平均售价"] = q;
                    }
                    else
                        dr["平均售价"] = 0;

                    dr["SKU名称"] = dtResultConnig.Rows[i]["SKU名称"];
                    dr["品牌"] = dtResultConnig.Rows[i]["品牌"];
                    dr["品类"] = dtResultConnig.Rows[i]["品类"];
                    dr["机型"] = dtResultConnig.Rows[i]["机型"];
                    dtResultSheetCount.Rows.Add(dr);
                }
            }


            #endregion
        }
    }
}
