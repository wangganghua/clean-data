using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace AVC_ClareData
{
    static class Program
    {
        /// <summary>
        /// 应用程序的主入口点。
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            //Application.Run(new DataClean());
            //Application.Run(new Form1());
            //Application.Run(new InsertXian());
            //Application.Run(new InsertXian_sjxh());
            Application.Run(new urlData());
        }
    }
}
