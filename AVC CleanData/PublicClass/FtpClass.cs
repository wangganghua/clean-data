using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Net;
using System.Windows.Forms;
using System.IO;
using System.Diagnostics;
using System.Threading;

namespace AVC_ClareData.PublicClass
{
    /// <summary>
    /// Frp服务器访问方法
    /// </summary>
  public  class FtpClass
    {
        /// <summary>
        /// Ftp访问请求
        /// </summary>
        FtpWebRequest request = null;
        /// <summary>
        /// Ftp服务器IP地址
        /// </summary>
        public string ftpServerIPAddress;
        /// <summary>
        /// Ftp服务器用户名
        /// </summary>
        string ftpServerUserID;
        /// <summary>
        /// Ftp服务器密码
        /// </summary>
        string ftpServerPassword;

        /// <summary>
        /// 有参构造函数
        /// </summary>
        /// <param name="ftpServerIPAddress">Ftp服务器IP地址，例如：192.168.1.1</param>
        /// <param name="ftpServerUserID">Ftp服务器用户名，例如：user001</param>
        /// <param name="ftpServerPassword">Ftp服务器密码，例如：pass</param>
        public FtpClass(string ftpServerIPAddress, string ftpServerUserID, string ftpServerPassword)
        {
            this.ftpServerIPAddress = ftpServerIPAddress;
            this.ftpServerUserID = ftpServerUserID;
            this.ftpServerPassword = ftpServerPassword;
        }

        /// <summary>
        /// 析构函数
        /// </summary>
        ~FtpClass()
        {

        }

        /// <summary>
        /// 创建FtpWebRequest
        /// </summary>
        /// <param name="path">文件夹名或文件名，如“Folder/File.txt”或“File.txt”</param>
        /// <param name="message">回传的信息</param>
        /// <returns></returns>
        private bool Connect(string path, ref string message)
        {
            try
            {
                //根据uri创建FtpWebRequest对象
                request = (FtpWebRequest)FtpWebRequest.Create(new Uri("ftp://" + ftpServerIPAddress + "/" + path));
                //设置与FTP服务器通信的凭据（用户名和口令）
                request.Credentials = new NetworkCredential(ftpServerUserID, ftpServerPassword);
                //指定数据传输类型
                request.UseBinary = true;
                request.UsePassive = false;
                request.KeepAlive = false;
                request.EnableSsl = false;
                request.Proxy = WebRequest.DefaultWebProxy;//设置不使用HTTP代理,如果出现出错,请注释掉本行
                return true;
            }
            catch (Exception ex)
            {
                message = "尝试连接路径或文件【" + path + "】的过程中发生错误，" + ex.Message;
                return false;
            }
        }

        ///<summary>
        ///上传文件到FTP服务器
        ///</summary>
        ///<param name="UploadFilePath">源文件路径，例如：D:\Template\a-模板.xls</param>
        ///<param name="folderName">目的地文件夹名称，例如：Template/常规报告/彩电/线上/周报</param>
        ///<param name="overwrite">重名是否覆盖</param>
        ///<param name="StatusLabel">显示信息的Label</param>
        ///<param name="progressBar">进度条</param>
        /// <param name="message">回传的信息</param>
        ///<returns></returns>
        public bool Upload(string UploadFilePath, string folderName, bool overwrite, Label StatusLabel, ProgressBar progressBar, ref string message)
        {
            if (!DirectoryExists(folderName, ref message))
                CreateDirectory(folderName, ref message);
            if (FileExists(folderName + "/" + Path.GetFileName(UploadFilePath), ref message))
            {
                if (Size(folderName + "/" + Path.GetFileName(UploadFilePath), ref message) == 0)
                    DeleteFile(folderName + "/" + Path.GetFileName(UploadFilePath), ref message);
                else
                {
                    if (overwrite)
                        DeleteFile(folderName + "/" + Path.GetFileName(UploadFilePath), ref message);
                    else
                        return true;
                }
            }
            FileInfo fileInfo = new FileInfo(UploadFilePath);
            //根据uri创建FtpWebRequest对象 
            Connect(folderName + "/" + Path.GetFileName(UploadFilePath), ref message);
            //上传文件时通知服务器文件的大小
            request.ContentLength = fileInfo.Length;
            //指定发送到FTP服务器的命令
            request.Method = WebRequestMethods.Ftp.UploadFile;
            //设置缓冲区域大小设置为2MB
            int buffLength = 2 * 1024 * 1024;
            byte[] buff = new byte[buffLength];
            int contentLength;
            //获取文件大小
            long FileSize = fileInfo.Length;
            //记录已上传大小
            long loaded = 0;
            //打开一个文件流 (System.IO.FileStream) 读取要上传的文件
            FileStream fileStream = fileInfo.OpenRead();
            //把上传的文件写入流
            Stream stream = request.GetRequestStream();
            try
            {
                //每次读文件流的2MB
                contentLength = fileStream.Read(buff, 0, buffLength);
                //判断流内容是否结束
                while (contentLength != 0)
                {
                    //把内容从File Stream写入Upload Stream
                    DateTime start = DateTime.Now;
                    stream.Write(buff, 0, contentLength);
                    loaded += contentLength;
                    DateTime end = DateTime.Now;
                    //计算速度，单位【字节/秒】
                    double TotalMilliseconds = (end - start).TotalMilliseconds;
                    if (TotalMilliseconds == 0)
                        TotalMilliseconds = 1;
                    double speed = (double)contentLength / TotalMilliseconds * 1000;
                    if (speed > 1024 * 1024)//适合转换为MB/s
                        StatusLabel.Invoke(new ThreadStart(delegate() { StatusLabel.Text = "正在上传文件【" + UploadFilePath + "】，上传速度" + Math.Round(speed / 1024 / 1024, 2) + "MB/s，剩余时间大约" + Math.Round((double)(FileSize - loaded) / speed, 0) + "秒，进度" + Math.Round(((float)loaded * 100 / FileSize), 2) + "%"; }));
                    else if (speed > 1024)//适合转换为KB/s
                        StatusLabel.Invoke(new ThreadStart(delegate() { StatusLabel.Text = "正在上传文件【" + UploadFilePath + "】，上传速度" + Math.Round(speed / 1024, 2) + "KB/s，剩余时间大约" + Math.Round((double)(FileSize - loaded) / speed, 0) + "秒，进度" + Math.Round(((float)loaded * 100 / FileSize), 2) + "%"; }));
                    else
                        StatusLabel.Invoke(new ThreadStart(delegate() { StatusLabel.Text = "正在上传文件【" + UploadFilePath + "】，上传速度" + Math.Round(speed, 2) + "字节/s，剩余时间大约" + Math.Round((double)(FileSize - loaded) / speed, 0) + "秒，进度" + Math.Round(((float)loaded * 100 / FileSize), 2) + "%"; }));
                    progressBar.Invoke(new ThreadStart(delegate() { progressBar.Value = (int)(loaded * 100 / FileSize); }));
                    //继续获取二进制流
                    contentLength = fileStream.Read(buff, 0, buffLength);
                }
                return true;
            }
            catch (Exception ex)
            {
                fileStream.Close();
                message = "在将文件【" + UploadFilePath + "】上传到【" + folderName + "】的过程中发生错误，" + ex.Message;
                return false;
            }
            finally
            {
                //关闭两个流
                stream.Close();
                fileStream.Close();
            }
        }

        ///<summary>
        ///上传文件到FTP服务器
        ///</summary>
        ///<param name="UploadFilePath">源文件路径，例如：D:\Template\a-模板.xls</param>
        ///<param name="folderName">目的地文件夹名称，例如：Template/常规报告/彩电/线上/周报</param>
        ///<param name="overwrite">重名是否覆盖</param>
        /// <param name="message">回传的信息</param>
        ///<returns></returns>
        public bool Upload(string UploadFilePath, string folderName, bool overwrite, ref string message)
        {
            if (!DirectoryExists(folderName, ref message))
                CreateDirectory(folderName, ref message);
            if (FileExists(folderName + "/" + Path.GetFileName(UploadFilePath), ref message))
            {
                if (Size(folderName + "/" + Path.GetFileName(UploadFilePath), ref message) == 0)
                    DeleteFile(folderName + "/" + Path.GetFileName(UploadFilePath), ref message);
                else
                {
                    if (overwrite)
                        DeleteFile(folderName + "/" + Path.GetFileName(UploadFilePath), ref message);
                    else
                        return true;
                }
            }
            FileInfo fileInfo = new FileInfo(UploadFilePath);
            //根据uri创建FtpWebRequest对象 
            Connect(folderName + "/" + Path.GetFileName(UploadFilePath), ref message);
            //上传文件时通知服务器文件的大小
            request.ContentLength = fileInfo.Length;
            //指定发送到FTP服务器的命令
            request.Method = WebRequestMethods.Ftp.UploadFile;
            //设置缓冲区域大小设置为2MB
            int buffLength = 2 * 1024 * 1024;
            byte[] buff = new byte[buffLength];
            int contentLength;
            //获取文件大小
            long FileSize = fileInfo.Length;
            //记录已上传大小
            long loaded = 0;
            //打开一个文件流 (System.IO.FileStream) 读取要上传的文件
            FileStream fileStream = fileInfo.OpenRead();
            //把上传的文件写入流
            Stream stream = request.GetRequestStream();
            try
            {
                //每次读文件流的2MB
                contentLength = fileStream.Read(buff, 0, buffLength);
                //判断流内容是否结束
                while (contentLength != 0)
                {
                    //把内容从File Stream写入Upload Stream
                    stream.Write(buff, 0, contentLength);
                    loaded += contentLength;
                    //继续获取二进制流
                    contentLength = fileStream.Read(buff, 0, buffLength);
                }
                return true;
            }
            catch (Exception ex)
            {
                fileStream.Close();
                message = "在将文件【" + UploadFilePath + "】上传到【" + folderName + "】的过程中发生错误，" + ex.Message;
                return false;
            }
            finally
            {
                //关闭两个流		
                stream.Close();
                fileStream.Close();
            }
        }

        /// <summary>
        /// 下载FTP服务器上的文件到指定的路径
        /// </summary>
        /// <param name="fileStorePath">下载的文件存储路径，不包括文件名和扩展名，例如：D:\Template</param>
        /// <param name="sourceFileName">源文件路径，例如：Template/常规报告/a.xls</param>
        /// <param name="overWritten">如已存在同名文件，指示是否覆盖</param>
        ///<param name="StatusLabel">显示信息的Label</param>
        ///<param name="progressBar">进度条</param>
        /// <param name="message">回传的信息</param>
        /// <returns></returns>
        public bool Download(string fileStorePath, string sourceFileName, bool overWritten, Label StatusLabel, ProgressBar progressBar, ref string message)
        {
            string newFileName = fileStorePath + "\\" + Path.GetFileName("ftp://" + ftpServerIPAddress + "/" + sourceFileName);
            if (!Directory.Exists(fileStorePath))
            {
                try
                {
                    Directory.CreateDirectory(fileStorePath);
                }
                catch (Exception ex)
                {
                    message = ex.Message;
                    return false;
                }
            }
            if (File.Exists(newFileName))
            {
                FileInfo info = new FileInfo(newFileName);
                if (info.Length == 0)
                    File.Delete(newFileName);
                else
                {
                    if (overWritten)
                        File.Delete(newFileName);
                    else
                        return true;
                }
            }
            Connect(sourceFileName, ref message);
            request.Method = WebRequestMethods.Ftp.DownloadFile;
            FtpWebResponse response = (FtpWebResponse)request.GetResponse();
            Stream ftpStream = response.GetResponseStream();
            FileStream outputStream = new FileStream(newFileName, FileMode.Create);
            try
            {
                //得到服务器文件流大小
                long ftpStreamSize = Size(sourceFileName, ref message);
                //设置缓冲为2MB
                int buffLength = 2 * 1024 * 1024;
                byte[] buff = new byte[buffLength];
                //获取文件大小
                long loaded = 0;
                //定义每次写入2MB
                int contentLength = ftpStream.Read(buff, 0, buffLength);
                while (contentLength > 0)
                {
                    //把内容写到本地 
                    DateTime start = DateTime.Now;
                    outputStream.Write(buff, 0, contentLength);
                    loaded += contentLength;
                    DateTime end = DateTime.Now;
                    //计算速度
                    double timediff = (end - start).TotalMilliseconds;
                    if (timediff == 0)
                        timediff = 1;
                    double speed = (double)contentLength / timediff * 1000;
                    if (speed > 1024 * 1024)//适合转换为MB/s
                        StatusLabel.Invoke(new ThreadStart(delegate() { StatusLabel.Text = "正在下载文件【" + sourceFileName + "】，下载速度" + Math.Round(speed / 1024 / 1024, 2) + "MB/s，剩余时间大约" + Math.Round((double)(ftpStreamSize - loaded) / speed, 0) + "秒，进度" + Math.Round(((float)loaded * 100 / ftpStreamSize), 2) + "%"; }));
                    else if (speed > 1024)//适合转换为KB/s
                        StatusLabel.Invoke(new ThreadStart(delegate() { StatusLabel.Text = "正在下载文件【" + sourceFileName + "】，下载速度" + Math.Round(speed / 1024, 2) + "KB/s，剩余时间大约" + Math.Round((double)(ftpStreamSize - loaded) / speed, 0) + "秒，进度" + Math.Round(((float)loaded * 100 / ftpStreamSize), 2) + "%"; }));
                    else
                        StatusLabel.Invoke(new ThreadStart(delegate() { StatusLabel.Text = "正在下载文件【" + sourceFileName + "】，下载速度" + Math.Round(speed, 2) + "字节/s，剩余时间大约" + Math.Round((double)(ftpStreamSize - loaded) / speed, 0) + "秒，进度" + Math.Round(((float)loaded * 100 / ftpStreamSize), 2) + "%"; }));
                    progressBar.Invoke(new ThreadStart(delegate() { progressBar.Value = (int)(loaded * 100 / ftpStreamSize); }));
                    //继续获取二进制流
                    contentLength = ftpStream.Read(buff, 0, buffLength);
                }
                return true;
            }
            catch (Exception ex)
            {
                message = "在将文件【" + sourceFileName + "】下载到【" + fileStorePath + "】的过程中发生错误，" + ex.Message;
                return false;
            }
            finally
            {
                ftpStream.Close();
                outputStream.Close();
                response.Close();
            }
        }

        /// <summary>
        /// 下载FTP服务器上的文件到指定的路径
        /// </summary>
        /// <param name="fileStorePath">下载的文件存储路径，不包括文件名和扩展名，例如：D:\Template</param>
        /// <param name="sourceFileName">源文件路径，例如：Template/常规报告/a.xls</param>
        /// <param name="overWritten">如已存在同名文件，指示是否覆盖</param>
        /// <param name="message">回传的信息</param>
        /// <returns></returns>
        public bool Download(string fileStorePath, string sourceFileName, bool overWritten, ref string message)
        {
            string newFileName = fileStorePath + "\\" + Path.GetFileName("ftp://" + ftpServerIPAddress + "/" + sourceFileName);
            if (!Directory.Exists(fileStorePath))
                Directory.CreateDirectory(fileStorePath);
            if (File.Exists(newFileName))
            {
                FileInfo info = new FileInfo(newFileName);
                if (info.Length == 0)
                    File.Delete(newFileName);
                else
                {
                    if (overWritten)
                        File.Delete(newFileName);
                    else
                        return true;
                }
            }
            Connect(sourceFileName, ref message);
            request.Method = WebRequestMethods.Ftp.DownloadFile;
            FtpWebResponse response = (FtpWebResponse)request.GetResponse();
            Stream ftpStream = response.GetResponseStream();
            FileStream outputStream = new FileStream(newFileName, FileMode.Create);
            try
            {
                //得到服务器文件流大小
                long ftpStreamSize = Size(sourceFileName, ref message);
                //设置缓冲为2MB
                int buffLength = 2 * 1024 * 1024;
                byte[] buff = new byte[buffLength];
                //获取文件大小
                long loaded = 0;
                //定义每次写入2MB
                int contentLength = ftpStream.Read(buff, 0, buffLength);
                while (contentLength > 0)
                {
                    //把内容写到本地 
                    outputStream.Write(buff, 0, contentLength);
                    loaded += contentLength;
                    //继续获取二进制流
                    contentLength = ftpStream.Read(buff, 0, buffLength);
                }
                return true;
            }
            catch (Exception ex)
            {
                message = "在将文件【" + sourceFileName + "】下载到【" + fileStorePath + "】的过程中发生错误，" + ex.Message;
                return false;
            }
            finally
            {
                ftpStream.Close();
                outputStream.Close();
                response.Close();
            }
        }

        /// <summary>
        /// 获取FTP服务器上指定路径下的文件列表，如果指定路径是文件夹，则返回该文件夹下的文件列表，否则返回元素个数为0的数组
        /// </summary>
        /// <param name="WebRequestMethods">WebRequestMethods.Ftp类型</param>
        /// <param name="path">文件夹名、根目录、文件名，根目录或文件时path为空，例如：Template/常规报告</param>
        /// <param name="message">回传的信息</param>
        /// <returns></returns>
        public string[] FileList(string WebRequestMethods, string path, ref string message)
        {
            try
            {
                if (!Connect(path, ref message))
                    return null;
                StringBuilder stringBuilder = new StringBuilder();
                request.Method = WebRequestMethods;
                WebResponse response = request.GetResponse();
                StreamReader streamReader = new StreamReader(response.GetResponseStream(), System.Text.Encoding.UTF8);//中文文件名
                string line = streamReader.ReadLine();
                while (line != null)
                {
                    stringBuilder.Append(line);
                    stringBuilder.Append("\n");
                    line = streamReader.ReadLine();
                }
                streamReader.Close();
                response.Close();
                return stringBuilder.ToString().Split(new char[] { '\n' }, StringSplitOptions.RemoveEmptyEntries);
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.Message);
                return null;
            }
        }

        /// <summary>
        /// 判断文件夹
        /// </summary>
        /// <param name="path">文件夹路径，例如：Template/常规报告</param>
        /// <param name="message">回传的信息</param>
        /// <returns></returns>
        public bool IsDirectory(string path, ref string message)
        {
            try
            {
                //若参数以字符【/】结尾，则去除此字符
                int end = path.LastIndexOf('/');
                if (end == path.Length - 1)
                {
                    path = path.Substring(0, path.Length - 1);
                    end = path.LastIndexOf('/');
                }
                //获取文件夹名（也可能是文件名）
                string ftpFileName = path.Substring(end + 1, path.Length - end - 1);
                //获取文件夹的上层目录详细信息
                if (end == -1)
                    path = "";
                else
                    path = path.Substring(0, end);

                string[] rootDetail = FileList(WebRequestMethods.Ftp.ListDirectoryDetails, path, ref message);
                for (int i = 0; i < rootDetail.Length; i++)
                    if (rootDetail[i].Contains("<DIR>") && rootDetail[i].EndsWith(ftpFileName))
                        return true;
                return false;
            }
            catch (Exception ex)
            {
                message = "尝试判断路径或文件【" + path + "】的具体类型时发生错误，" + ex.Message;
                return false;
            }
        }

        /// <summary>
        /// 判断文件
        /// </summary>
        /// <param name="path">文件路径，例如：Template/常规报告/a.xls</param>
        /// <param name="message">回传的信息</param>
        /// <returns></returns>
        public bool IsFile(string path, ref string message)
        {
            try
            {
                //若参数以字符【/】结尾，则去除此字符
                int end = path.LastIndexOf('/');
                if (end == path.Length - 1)
                {
                    path = path.Substring(0, path.Length - 1);
                    end = path.LastIndexOf('/');
                }
                //获取文件名（也可能是文件夹名）
                string ftpFileName = path.Substring(end + 1, path.Length - end - 1);
                //获取文件的上层目录详细信息
                string[] rootDetail = FileList(WebRequestMethods.Ftp.ListDirectoryDetails, path.Substring(0, end), ref message);
                for (int i = 0; i < rootDetail.Length; i++)
                    if (!rootDetail[i].Contains("<DIR>") && rootDetail[i].EndsWith(ftpFileName))
                        return true;
                return false;
            }
            catch (Exception ex)
            {
                message = "尝试判断路径或文件【" + path + "】的具体类型时发生错误，" + ex.Message;
                return false;
            }
        }

        /// <summary>
        /// 删除FTP服务器上的文件
        /// </summary>
        /// <param name="fileName">源文件路径，例如：Template/常规报告/a.xls</param>
        /// <param name="message">回传的信息</param>
        /// <returns></returns>
        public bool DeleteFile(string fileName, ref string message)
        {
            try
            {
                if (!FileExists(fileName, ref message))
                    return true;
                if (!Connect(fileName, ref message))
                    return false;
                // 指定执行什么命令
                request.Method = WebRequestMethods.Ftp.DeleteFile;
                FtpWebResponse response = (FtpWebResponse)request.GetResponse();
                response.Close();
                return true;
            }
            catch (Exception ex)
            {
                message = "尝试删除文件【" + fileName + "】时发生错误，" + ex.Message;
                return false;
            }
        }

        /// <summary>
        /// 在FTP服务器上创建文件夹、目录
        /// </summary>
        /// <param name="dirName">文件夹名称，如“Folder1/Folder2”或“Folder3”</param>
        /// <param name="message">回传的信息</param>
        /// <returns></returns>
        public bool CreateDirectory(string dirName, ref string message)
        {
            try
            {
                if (dirName.LastIndexOf("/") > -1)
                {
                    if (!DirectoryExists(dirName.Remove(dirName.LastIndexOf("/")), ref message))
                        CreateDirectory(dirName.Remove(dirName.LastIndexOf("/")), ref message);
                }
                if (!Connect(dirName, ref message))
                    return false;
                request.Method = WebRequestMethods.Ftp.MakeDirectory;
                FtpWebResponse response = (FtpWebResponse)request.GetResponse();
                response.Close();
                return true;
            }
            catch (Exception ex)
            {
                message = "尝试创建路径【" + dirName + "】时发生错误，" + ex.Message;
                return false;
            }
        }

        /// <summary>
        /// 在FTP服务器上删除文件夹、目录
        /// </summary>
        /// <param name="dirName">文件夹名称，如“Folder1/Folder2”或“Folder3”</param>
        /// <param name="message">回传的信息</param>
        /// <returns></returns>
        public bool DeleteDirectory(string dirName, ref string message)
        {
            try
            {
                if (!Connect(dirName, ref message))
                    return false;
                request.Method = WebRequestMethods.Ftp.RemoveDirectory;
                FtpWebResponse response = (FtpWebResponse)request.GetResponse();
                response.Close();
                return true;
            }
            catch (Exception ex)
            {
                message = "尝试删除路径【" + dirName + "】时发生错误，" + ex.Message;
                return false;
            }
        }

        /// <summary>
        /// 获得FTP服务器上某个文件的大小
        /// </summary>
        /// <param name="fileName">文件路径，例如：Template/常规报告/a.xls</param>
        /// <param name="message">回传的信息</param>
        /// <returns></returns>
        public long Size(string fileName, ref string message)
        {
            long fileSize = 0;
            try
            {
                FileInfo fileInf = new FileInfo(fileName);
                if (!Connect(fileName, ref message))
                    return 0;
                request.Method = WebRequestMethods.Ftp.GetFileSize;
                FtpWebResponse response = (FtpWebResponse)request.GetResponse();
                fileSize = response.ContentLength;
                response.Close();
                return fileSize;
            }
            catch (Exception ex)
            {
                message = "尝试获取文件【" + fileName + "】的大小时发生错误，" + ex.Message;
                return 0;
            }
        }

        /// <summary>
        /// 重命名FTP服务器上的某个文件
        /// </summary>
        /// <param name="currentFilename">原文件名，例如：Template/常规报告/a.xls</param>
        /// <param name="newFileName">重命名后的文件名，例如：Template/常规报告/b.xls</param> 
        /// <param name="message">回传的信息</param>
        /// <returns></returns>
        public bool Rename(string currentFilename, string newFileName, ref string message)
        {
            try
            {
                FileInfo fileInf = new FileInfo(currentFilename);
                string uri = "ftp://" + ftpServerIPAddress + "/" + fileInf.Name;
                if (!Connect(currentFilename, ref message))
                    return false;
                request.Method = WebRequestMethods.Ftp.Rename;
                request.RenameTo = newFileName;
                FtpWebResponse response = (FtpWebResponse)request.GetResponse();
                response.Close();
                return true;
            }
            catch (Exception ex)
            {
                message = "尝试重命名文件【" + currentFilename + "】到【" + newFileName + "】时发生错误，" + ex.Message;
                return false;
            }
        }

        /// <summary>
        /// 判断文件是否存在
        /// </summary>
        /// <param name="path">文件路径，例如：Template/常规报告/a.xls</param>
        /// <param name="message">回传的信息</param>
        /// <returns></returns>
        public bool FileExists(string path, ref string message)
        {
            string[] list = FileList(WebRequestMethods.Ftp.ListDirectory, path, ref message);
            if (IsFile(path, ref message) && list.Length != 0)
                return true;
            return false;
        }

        /// <summary>
        /// 判断文件夹是否存在
        /// </summary>
        /// <param name="path">文件夹路径，例如：Template/常规报告</param>
        /// <param name="message">回传的信息</param>
        /// <returns></returns>
        public bool DirectoryExists(string path, ref string message)
        {
            string[] list = FileList(WebRequestMethods.Ftp.ListDirectory, path, ref message);
            if (list == null)
                return false;
            if (IsDirectory(path, ref message))
                return true;
            return false;
        }

        /// <summary>
        /// 获取指定文件的创建日期
        /// </summary>
        /// <param name="path">Ftp文件路径</param>
        /// <param name="message"></param>
        /// <returns></returns>
        public DateTime FileCreateTime(string path, ref string message)
        {
            DateTime createDateTime = DateTime.MinValue;
            if (IsFile(path, ref message))
            {
                string[] details = FileList(WebRequestMethods.Ftp.ListDirectoryDetails, path, ref message);
                if (details != null && details.Length != 0)
                {
                    string[] createTime = details[0].Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
                    if (createTime.Length == 4)
                    {
                        string[] mdy = createTime[0].Split('-');
                        createDateTime = Convert.ToDateTime(mdy[2] + "-" + mdy[0] + "-" + mdy[1] + " " + createTime[1]);
                    }
                }
            }
            return createDateTime;
        }
    }
}

