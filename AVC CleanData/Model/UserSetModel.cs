using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace AVC_ClareData.Model
{
   public class UserSetModel
    {
       /// <summary>
       /// ftp服务器ip
       /// </summary>
        private string ftpiIp;
        public string FtpIp
        {
            get { return ftpiIp = "192.168.2.236"; }
            set { ftpiIp = value; }
        }
       /// <summary>
       /// ftp登陆用户名
       /// </summary>
        private string ftpUserName;
        public string FtpUserName
        {
            get { return ftpUserName = "Stuart"; }
            set { ftpUserName = value; }
        }
       /// <summary>
       /// ftp登陆用户密码
       /// </summary>
        private string ftpUserPwd;
        public string FtpUserPwd
        {
            get { return ftpUserPwd = "jetaime001@"; }
            set { ftpUserPwd = value; }
        }
       
    }
}
