using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

using Tamir.SharpSsh.jsch;
using Tamir.SharpSsh.jsch.examples;
using System.Diagnostics;
using System.IO;

namespace AVC_ClareData
{
    public partial class Form1 : Form
    {
        MySqlConnection mysql = new MySqlConnection();
        public Form1()
        {
            InitializeComponent();
        }
        private void OpenSSH()
        {
            string host = "1.119.7.235";
            string user = "root";
            string pwd = "";
            //SshShell shell = new SshShell(host, user);
            //ssh
            JSch jsch = new JSch();
            //String file = InputForm.GetFileFromUser(@"E:\A工作文件\AVC CleanData\AVC CleanData\id_rsa_2048.2048");
            string file = @"E:\A工作文件\AVC CleanData\AVC CleanData\id_rsa_2048.2048";
            jsch.addIdentity(file);

            Session session = jsch.getSession(user, host);

            UserInfo ui = new MyUserInfo();
            session.setUserInfo(ui);
            try
            {
                session.connect();
                // Channel channel = session.openChannel("shell");
            }
            catch (Exception e)
            {
                Debug.WriteLine(e);
            }
            finally
            {
                session.disconnect();
            }
        }
        private string[] weeks = new string[] { };
        string yearweek = "18W02";
        private void button1_Click(object sender, EventArgs e)
        {
            //OpenSSH();
            string Wdd = string.Empty;
            weeks = new string[9];
            if (yearweek.ToUpper().Contains("W53"))
            {
              
                weeks = PreviousWeeks("17W53", 9);
                //Debug.WriteLine(weeks[0]);
                //Wdd = DateRangeOfWeek(yearweek);
            }
            weeks = PreviousWeeks(yearweek, 101);

            Wdd = DateRangeOfWeek(yearweek);
            string[] allWeeks = new string[] { };
            string[] yearallWeeks = new string[] { };
            string[] lastyearallWeeks = new string[] { };
            yearallWeeks = PreviousWeeks(yearweek, Int32.Parse(yearweek.Substring(3, 2)));
            lastyearallWeeks = PreviousWeeks((Int32.Parse(yearweek.Substring(0, 2)) - 1 + "W52"), 52);
            Array.Reverse(yearallWeeks);
            Array.Reverse(lastyearallWeeks);
            allWeeks = yearallWeeks.Concat(lastyearallWeeks).ToArray();

            }

        /// <summary>
        /// 获得前n个星期的标准格式
        /// </summary>
        /// <param name="week">周度，表示方式如12W10表示2012年第10周，每年第一周如不为星期一，则将上一年剩余的不足一周的天数并入本年的第一周，如果这样算的话超过53周，则将上述第一周忽略</param>
        /// <param name="n">返回周度的数目</param>
        /// <returns>包含参数Week在内的之前n个周</returns>
        public static string[] PreviousWeeks(string week, int n)
        {
            if (n == 0)
                return new string[0];
            //一个星期中每一天的名称
            string[] DayName = { "monday", "tuesday", "wednesday", "thursday", "friday", "saturday", "sunday" };
            string[] TW = new string[n];
            int Year = Int32.Parse(week.Substring(0, 2)) + 2000;
            int WeekOfYear = Int32.Parse(week.Substring(3, 2));
            int LastYear = Year - 1;
            int MaxWeekValue = 0;
            DateTime dt = Convert.ToDateTime(LastYear.ToString() + "-12-31");
            for (int i = 0; i < 7; i++)
                if (dt.DayOfWeek.ToString().ToLower() == DayName[i])
                    if (i != 6)
                        dt = Convert.ToDateTime(LastYear + "-12-" + (30 - i).ToString());
            if (dt.DayOfYear % 7 != 0)
                MaxWeekValue = dt.DayOfYear / 7 + 1;
            else
                MaxWeekValue = dt.DayOfYear / 7;
            if (MaxWeekValue > 52)
                //if (LastYear == 2017)
                //    MaxWeekValue = 52;
                //else
                    MaxWeekValue = 53;
            if (WeekOfYear > n - 1)
            {
                for (int i = 0; i < n; i++)
                    if (WeekOfYear + i - n < 9)
                        TW[i] = (Year - 2000).ToString() + "W0" + (WeekOfYear + i - n + 1).ToString();
                    else
                        TW[i] = (Year - 2000).ToString() + "W" + (WeekOfYear + i - n + 1).ToString();
            }
            else
            {
                for (int i = 0; i < n - WeekOfYear; i++)
                {
                    if (MaxWeekValue + WeekOfYear + i - n + 1 < 10)
                        TW[i] = ((Year - 2000) - 1).ToString() + "W0" + (MaxWeekValue + WeekOfYear + i - n + 1).ToString();
                    else
                        TW[i] = ((Year - 2000) - 1).ToString() + "W" + (MaxWeekValue + WeekOfYear + i - n + 1).ToString();
                }
                for (int i = n - WeekOfYear; i < n; i++)
                    if (i + WeekOfYear - n < 9)
                        TW[i] = (Year - 2000).ToString() + "W0" + (i + WeekOfYear - n + 1).ToString();
                    else
                        TW[i] = (Year - 2000).ToString() + "W" + (i + WeekOfYear - n + 1).ToString();
            }
            return TW;
        }

        public static string DateRangeOfWeek(string week)
        {
            if (week.Length != 5)
                throw new ArgumentException("输入周度格式不正确！");
            if (week[2]!= 'W')
                throw new ArgumentException("输入周度格式不正确！");
            int check;
            if (!Int32.TryParse(week.Remove(2, 1), out check))
                throw new ArgumentException("输入周度格式不正确！");
            //一个星期中每一天的名称
            string[] DayName = { "monday", "tuesday", "wednesday", "thursday", "friday", "saturday", "sunday" };
            //非闰年时每个月的天数
            int[] DaysOfEveryMonth = { 31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31 };
            int day = 0;
            string[] DateRange = new string[2];
            string[] YearAndWeek = week.Split('W');
            string str = "20" + YearAndWeek[0] + "-01-01";
            DateTime date = Convert.ToDateTime(str);
            int y = Convert.ToInt32("20" + YearAndWeek[0]);
            int w = Convert.ToInt32(YearAndWeek[1]);
            if (y == 2012 || y == 2017)
                w++;
            for (int j = 0; j < 7; j++)
                if (date.DayOfWeek.ToString().ToLower().Equals(DayName[j]))
                {
                    day = j + 1;
                    break;
                }
            int daycount = w * 7 -5- day;//daycount表示y年的第w周的星期一是该年中的第daycount天
            if (DateTime.IsLeapYear(y))//判断闰年，闰年2月为29天
                DaysOfEveryMonth[1] = 29;
            else
                DaysOfEveryMonth[1] = 28;
            int m = 1;
            for (int i = 0; i < DaysOfEveryMonth.Length; i++)
            {
                if (daycount > DaysOfEveryMonth[m - 1])
                {
                    daycount -= DaysOfEveryMonth[m - 1];
                    m++;
                }
            }
            int d = daycount;
            if (d > 0)
            {
                if (m < 10)
                    DateRange[0] += "0" + m.ToString();
                else
                    DateRange[0] += m.ToString();
                if (d < 10)
                    DateRange[0] += "0" + d.ToString();
                else
                    DateRange[0] += d.ToString();
                if (m < 12)
                {
                    if (d + 6 > DaysOfEveryMonth[m - 1])
                    {
                        if (m + 1 < 10)
                            DateRange[1] += "0" + (m + 1).ToString();
                        else
                            DateRange[1] += (m + 1).ToString();
                        if (d + 6 - DaysOfEveryMonth[m - 1] < 10)
                            DateRange[1] += "0" + (d + 6 - DaysOfEveryMonth[m - 1]).ToString();
                        else
                            DateRange[1] += (d + 6 - DaysOfEveryMonth[m - 1]).ToString();
                    }
                    else
                    {
                        if (m < 10)
                            DateRange[1] += "0" + m.ToString();
                        else
                            DateRange[1] += m.ToString();
                        if (d + 6 < 10)
                            DateRange[1] += "0" + (d + 6).ToString();
                        else
                            DateRange[1] += (d + 6).ToString();
                    }
                }
                else
                {
                    if (d + 6 > 31)
                    {
                        DateRange[1] += "01";
                        //DateRange[1] += (d - 24).ToString();
                        if ((d - 25) < 10)
                            DateRange[1] += "0" + (d - 25).ToString();
                        else
                            DateRange[1] += (d - 25).ToString();//不知道什么情况
                    }
                    else
                    {
                        DateRange[1] += "12";
                        if (d + 6 > 9)
                            DateRange[1] += (d + 6).ToString();
                        else
                            DateRange[1] += "0" + (d + 6).ToString();
                    }
                }
            }
            else
            {
                DateRange[0] = "12" + (33 - day).ToString();
                DateRange[1] = "010" + (8 - day).ToString();
            }
            return DateRange[0].Substring(0, 2) + "." + DateRange[0].Substring(2, 2) + "-" + DateRange[1].Substring(0, 2) + "." + DateRange[1].Substring(2, 2);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            string url = "http://detail.tmall.com/item.htm?id=10032966948&amp;areaid=&amp;";
            WebBrowser webBrowser = new WebBrowser();  // 创建一个WebBrowser
            webBrowser.ScrollBarsEnabled = false;  // 隐藏滚动条
            webBrowser.Navigate(url);  // 打开网页
            webBrowser.DocumentCompleted += new WebBrowserDocumentCompletedEventHandler(webBrowser_DocumentCompleted);  // 增加网页加载完成事件处理函数
        }

        /// <summary>
        /// 网页加载完成事件处理函数
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        void webBrowser_DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
            WebBrowser webBrowser = (WebBrowser)sender;

            // 网页加载完毕才保存
            if (webBrowser.ReadyState == WebBrowserReadyState.Complete)
            {
                // 获取网页高度和宽度,也可以自己设置
                //int height = webBrowser.Document.Body.ScrollRectangle.Height;
                //int width = webBrowser.Document.Body.ScrollRectangle.Width;
                int height = 700;
                int width = 1050;
                HtmlElement em = webBrowser.Document.GetElementById("J_1730_110100");
                Debug.WriteLine(em);
                // 调节webBrowser的高度和宽度
                webBrowser.Height = height;
                webBrowser.Width = width;

                Bitmap bitmap = new Bitmap(width, height);  // 创建高度和宽度与网页相同的图片
                Rectangle rectangle = new Rectangle(0, 0, width, height);  // 绘图区域
                webBrowser.DrawToBitmap(bitmap, rectangle);  // 截图

                // 保存图片对话框
                SaveFileDialog saveFileDialog = new SaveFileDialog();
                saveFileDialog.Filter = "JPEG (*.jpg)|*.jpg|PNG (*.png)|*.png";
                saveFileDialog.ShowDialog();

                bitmap.Save(saveFileDialog.FileName);  // 保存图片
            }
        }


        private void button4_Click(object sender, EventArgs e)
        {
            test();
        }

        private void test()
        {
            string sql = "SELECT DISTINCT 输出字段名,columnname='分段表' FROM 数据导出_分段设置表 WHERE 品类='干衣机' " +
                "UNION ALL " +
                "SELECT DISTINCT 字段名称,columnname='属性表' FROM 数据导出_属性字段设置表 WHERE 品类='榨汁机' " +
                "UNION ALL " +
                "SELECT DISTINCT 字段名称,columnname='商场表' FROM 数据导出_商场字段设置表 WHERE 品类='榨汁机' AND 字段名称!='连锁名称'";
            DataTable dtx = mysql.GetdtTable(sql);
            if (treeView1.Nodes.Count > 0)
            {
                treeView1.Nodes.Clear();
            }
            BindTreeViewFatherNodeData(dtx, "columnname",treeView1);
        }
        private void BindTreeViewFatherNodeData(DataTable dtTreeData, string fatherColumnName,TreeView tv)
        {
            if (dtTreeData.Rows.Count > 0)
            {
                var columnName = from p in dtTreeData.AsEnumerable()
                                 group p by new { column = p.Field<object>(fatherColumnName) } into g
                                 select g.Key.column;
                if (columnName.Count() > 0)
                {
                    foreach (var colu in columnName)
                    {
                        if (colu.ToString() == "商场表")
                            continue;
                        Debug.WriteLine(colu);
                        TreeNode tn = new TreeNode();
                        tn.Text = colu.ToString();
                        tn.Checked = true;
                        treeView1.Nodes.Add(tn);
                        //绑定子节点
                        BindTreeViewChildrenNodeData(dtTreeData, "输出字段名", fatherColumnName, tn.Text, tn);
                    }
                }
            }
        }

        private void BindTreeViewChildrenNodeData(DataTable dtTreeData, string childrenColumnName, string fathercolumnName, string fathercolumnValue, TreeNode tvn)
        {
            if (dtTreeData.Rows.Count > 0)
            {
                var columnName = from p in dtTreeData.AsEnumerable()
                                 where p.Field<object>(fathercolumnName).ToString().Trim() == fathercolumnValue
                                 group p by new { column = p.Field<object>(childrenColumnName) } into g
                                 select g.Key.column;
                if (columnName.Count() > 0)
                {
                    foreach (var colu in columnName)
                    {
                        Debug.WriteLine(colu);
                        TreeNode tn = new TreeNode();
                        tn.Text = colu.ToString();
                        tn.Checked = true;
                     
                        tvn.Nodes.Add(tn);
                        tvn.ExpandAll();//展开节点
                       
                    }
                }
            }
        }

        //递归子节点跟随其全选或全不选
        private void ChangeChild(TreeNode node, bool state)
        {
            node.Checked = state;
            foreach (TreeNode tn in node.Nodes)
                ChangeChild(tn, state);
        }
        //递归父节点跟随其全选或全不选
        private void ChangeParent(TreeNode node)
        {
            if (node.Parent != null)
            {
                //兄弟节点被选中的个数
                int brotherNodeCheckedCount = 0;
                //遍历该节点的兄弟节点
                foreach (TreeNode tn in node.Parent.Nodes)
                {
                    if (tn.Checked == true)
                        brotherNodeCheckedCount++;
                }
                //兄弟节点全没选，其父节点也不选
                if (brotherNodeCheckedCount == 0)
                {
                    node.Parent.Checked = false;
                    ChangeParent(node.Parent);
                }
                //兄弟节点只要有一个被选，其父节点也被选
                if (brotherNodeCheckedCount >= 1)
                {
                    node.Parent.Checked = true;
                    ChangeParent(node.Parent);
                }
            }
        }

        //private void treeView1_MouseClick(object sender, MouseEventArgs e)
        //{
        //    bool isblack = false;
        //    TreeNode node = treeView1.GetNodeAt(new Point(e.X, e.Y));
        //    if (node != null)
        //    {
        //        if (node.Bounds.Contains(e.X, e.Y) == false)
        //            isblack = true;
        //        //ChangeChild(node, node.Checked);//影响子节点
        //        //ChangeParent(node);//影响父节点
        //        ccd(node, node.Checked);
        //        if (node.Parent != null)
        //            fcd(node, node.Checked);
        //    }
        //    else
        //        isblack = true;
        //}
        private void ccd(TreeNode node, bool istrue)
        {
            node.Checked = istrue;
            foreach (TreeNode nd in node.Nodes)
            {
                nd.Checked = istrue;
            }
        }
        private void fcd(TreeNode node, bool istrue)
        {
            node.Checked = istrue;
            if (node.Parent != null)
                node.Parent.Checked = istrue;
        }

        private void getTree()
        {
            if (treeView1.Nodes.Count > 0)
            {
                foreach (TreeNode tn in treeView1.Nodes)
                {
                    if (tn.Checked == false)
                        continue;
                    Debug.WriteLine(tn.Text);
                    if (tn.Nodes.Count > 0)
                    {
                        foreach (TreeNode ctn in tn.Nodes)
                        {
                            if (ctn.Checked == false)
                                continue;
                            Debug.WriteLine(ctn.Text);
                        }
                    }
                }
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            //getTree();
            List<string> aaa = getTreeNodes(treeView1, "商场表");
            Debug.WriteLine("ddd");
        }

        private List<string> getTreeNodes(TreeView tv, string tvfatherNodeName)
        {
            if (tv.Nodes.Count > 0)
            {
                foreach (TreeNode tn in tv.Nodes)
                {
                    if (tn.Checked == false)
                        continue;
                    if (tn.Text.ToString().Trim() != tvfatherNodeName.ToString().Trim())
                        continue;
                    List<string> childNodes = new List<string>();
                    if (tn.Nodes.Count > 0)
                    {
                        foreach (TreeNode ctn in tn.Nodes)
                        {
                            if (ctn.Checked == false)
                                continue;
                            childNodes.Add(ctn.Text);
                        }
                    }
                    return childNodes;
                }
            }
            return null;
        }

        private void treeView1_AfterCheck(object sender, TreeViewEventArgs e)
        {
            int level = -1;
            TreeNode node = e.Node;
            if (e.Node.Level >= 0)
            {
                level = e.Node.Level;
                if (level == 0)
                {
                    if (e.Node.Checked == true)
                    {

                    }
                }
            }
            else
                level = -1;
            if (level == -1)
            {
                Debug.WriteLine(e.Node.Level);
            }
        }
        

        #region check选择事件

        private bool nextCheck(TreeNode n)   //判断同级的节点是否全选
        {
            foreach (TreeNode tn in n.Parent.Nodes)
            {
                if (tn.Checked == false) return false;
            }
            return true;
        }

        private bool nextNotCheck(TreeNode n)  //判断同级的节点是否全不选
        {
            if (n.Checked == true)
            {
                return false;
            }
            if (n.NextNode == null)
            {
                return true;
            }

            return this.nextNotCheck(n.NextNode);
        }

        private void cycleChild(TreeNode tn, bool check)    //遍历节点下的子节点
        {
            if (tn.Nodes.Count != 0)
            {
                foreach (TreeNode child in tn.Nodes)
                {
                    child.Checked = check;
                    if (child.Nodes.Count != 0)
                    {
                        cycleChild(child, check);
                    }
                }
            }
            else
                return;
        }

        private void cycleParent(TreeNode tn, bool check)    //遍历节点上的父节点
        {
            if (tn.Parent != null)
            {
                if (nextCheck(tn))
                {
                    tn.Parent.Checked = true;
                }
                else
                {
                    tn.Parent.Checked = false;
                }
                cycleParent(tn.Parent, check);
            }
            return;
        }

        //     afterCheck
        private void treeViewTest_AfterCheck(object sender, TreeViewEventArgs e)    //当选中或取消选中树节点上的复选框时发生
        {

        }

        #endregion
    }
    /// <summary>
    /// A user info for getting user data
    /// </summary>
    public class MyUserInfo : UserInfo
    {
        /// <summary>
        /// Holds the key file passphrase
        /// </summary>
        private String passphrase;

        /// <summary>
        /// Returns the user password
        /// </summary>
        public String getPassword() { return null; }

        /// <summary>
        /// Prompt the user for a Yes/No input
        /// </summary>
        public bool promptYesNo(String str)
        {
            return InputForm.PromptYesNo(str);
        }

        /// <summary>
        /// Returns the user passphrase (passwd for the private key file)
        /// </summary>
        public String getPassphrase() { return passphrase; }

        /// <summary>
        /// Prompt the user for a passphrase (passwd for the private key file)
        /// </summary>
        public bool promptPassphrase(String message)
        {
            passphrase = InputForm.GetUserInput(message, true);
            return true;
        }

        /// <summary>
        /// Prompt the user for a password
        /// </summary>
        public bool promptPassword(String message) { return true; }
        public void showMessage(String message)
        {
            InputForm.ShowMessage(message);
        }
             
    }
}
