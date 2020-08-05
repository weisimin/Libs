using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net;
using System.Windows.Forms;
using System.Xml;
using System.Web.Security;
using System.Text.RegularExpressions;
using System.Data.SqlClient;
using System.Configuration;
using System.Data;
using System.IO;
using System.IO.Compression;

using System.Web.Configuration;
using System.Runtime.InteropServices;
using System.Web.Script.Serialization;
using System.Collections;

namespace NetFramework
    {
        #region 用户检验相关
        public class Util_User
        {

            public static bool? ValidateUser(string UserName, string Password)
            {
                if (Membership.GetUser(UserName) == null)
                {
                    return null;
                }
                return Membership.ValidateUser(UserName, Password);

            }

            //public static bool? ValidateWebUser(string UserName, string Password, ref Guid ProviderUserKey, ref string AspxAuth, ref CookieContainer otscookie)
            //{
            //    WeixinRoboot.RobootWeb.WebService ws = new WeixinRoboot.RobootWeb.WebService();
            //    ws.CookieContainer = new CookieContainer();

            //    string Result = ws.UserLogIn(UserName, Password);
            //    if (Result.Contains("错误"))
            //    {
            //        return null;
            //    }

            //    else
            //    {
            //        ProviderUserKey = Guid.Parse(Result);
            //        AspxAuth = ws.GetUserToken(UserName, Password);
            //        otscookie = ws.CookieContainer;
            //        return true;
            //    }
            //}

        }
        #endregion

        #region  邮件相关
        public class Util_Email
        {
            public static void EMail_SendEmail(string Server, Int32 Port, bool EnableSSL, string UserName, string Password, System.Net.Mail.MailAddress FromAddress, List<System.Net.Mail.MailAddress> To, List<System.Net.Mail.MailAddress> CC, string Subject, string BodyHtml, List<System.Net.Mail.Attachment> attachlist)
            {
                //简单邮件传输协议类
                System.Net.Mail.SmtpClient client = new System.Net.Mail.SmtpClient();
                client.Host = Server;//邮件服务器
                client.Port = Port;//smtp主机上的端口号,默认是25.
                client.EnableSsl = EnableSSL;

                client.DeliveryMethod = System.Net.Mail.SmtpDeliveryMethod.Network;//邮件发送方式:通过网络发送到SMTP服务器

                client.UseDefaultCredentials = true;
                client.Credentials = new NetworkCredential(UserName, Password); ;//凭证,发件人登录邮箱的用户名和密码



                //电子邮件信息类
                System.Net.Mail.MailMessage mailMessage = new System.Net.Mail.MailMessage();//创建一个电子邮件类
                                                                                            //似乎部分邮件不允许显示人改名
                mailMessage.From = FromAddress;

                foreach (System.Net.Mail.MailAddress item in To)
                {
                    mailMessage.To.Add(item);
                }
                foreach (System.Net.Mail.MailAddress item in CC)
                {
                    mailMessage.CC.Add(item);
                }


                mailMessage.Subject = Subject;
                mailMessage.SubjectEncoding = System.Text.Encoding.UTF8;//邮件主题编码

                mailMessage.Body = BodyHtml;//可为html格式文本
                mailMessage.BodyEncoding = System.Text.Encoding.GetEncoding("UTF-8");//邮件内容编码
                mailMessage.IsBodyHtml = true;//邮件内容是否为html格式

                mailMessage.Priority = System.Net.Mail.MailPriority.High;//邮件的优先级,有三个值:高(在邮件主题前有一个红色感叹号,表示紧急),低(在邮件主题前有一个蓝色向下箭头,表示缓慢),正常(无显示).
                                                                         //附件
                foreach (System.Net.Mail.Attachment att in attachlist)
                {
                    mailMessage.Attachments.Add(att);
                }
                //异步传输事件
                client.Timeout = 60000;
                client.SendCompleted += new System.Net.Mail.SendCompletedEventHandler(client_SendCompleted);
                try
                {

                    client.SendAsync(mailMessage, mailMessage);//发送邮件
                }
                catch (Exception AnyError)
                {
                    MessageBox.Show(AnyError.Message);
                }


            }

            static void client_SendCompleted(object sender, System.ComponentModel.AsyncCompletedEventArgs e)
            {
                if (e.Error != null)
                {
                    MessageBox.Show(e.Error.Message);
                }
                else
                {
                    MessageBox.Show("发送完成");
                }
            }

            /// <summary>
            /// 验证EMail是否合法
            /// </summary>
            /// <param name="email">要验证的Email</param>
            public static bool IsEmail(string emailStr)
            {
                return Regex.IsMatch(emailStr, @"^([\w-\.]+)@((\[[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.)|(([\w-]+\.)+))([a-zA-Z]{2,4}|[0-9]{1,3})(\]?)$");
            }

            public static string SplitGetLast(string EMLFolder)
            {
                string Result = "";
                string[] FullList = EMLFolder.Split("\"\"".ToCharArray(), StringSplitOptions.RemoveEmptyEntries);
                for (int i = FullList.Length - 1; i >= 0; i--)
                {
                    if ((FullList[i] != ""))
                    {
                        Result = FullList[i];
                        break;
                    }
                }
                return Result;
            }

        }
        #endregion

        #region 半数据转换
        public class Util_Convert
        {
            public static bool HalfBool(string Value)
            {
                try
                {
                    return Convert.ToBoolean(Value);
                }
                catch (Exception)
                {

                    return false;
                }
            }
            public static string ToString(object param)
            {
                if (param == null)
                {
                    return "";
                }
                else
                {
                    return param.ToString();
                }
            }
        }
        #endregion

        #region
        //Quoted-Printable 解码
        public class Util_Quoted
        {
            private const string QpSinglePattern = "(\\=([0-9A-F][0-9A-F]))";

            private const string QpMutiplePattern = @"((\=[0-9A-F][0-9A-F])+=?\s*)+";

            public static string Decode(string contents, Encoding encoding)
            {
                if (contents == null)
                {
                    throw new ArgumentNullException("contents");
                }

                // 替换被编码的内容
                string result = Regex.Replace(contents, QpMutiplePattern, new MatchEvaluator(delegate (Match m)
                {
                    List<byte> buffer = new List<byte>();
                    // 把匹配得到的多行内容逐个匹配得到后转换成byte数组
                    MatchCollection matches = Regex.Matches(m.Value, QpSinglePattern, RegexOptions.IgnoreCase | RegexOptions.Compiled);
                    foreach (Match match in matches)
                    {
                        buffer.Add((byte)HexToByte(match.Groups[2].Value.Trim()));
                    }
                    return encoding.GetString(buffer.ToArray());
                }), RegexOptions.IgnoreCase | RegexOptions.Compiled);

                // 替换多余的链接=号
                result = Regex.Replace(result, @"=\s+", "");

                return result;
            }

            private static int HexToByte(string hex)
            {
                int num1 = 0;
                string text1 = "0123456789ABCDEF";
                for (int num2 = 0; num2 < hex.Length; num2++)
                {
                    if (text1.IndexOf(hex[num2]) == -1)
                    {
                        return -1;
                    }
                    num1 = (num1 * 0x10) + text1.IndexOf(hex[num2]);
                }
                return num1;
            }

        }

        #endregion

        public class Util_Sql
        {

            public static object RunSqlText(string ConnectionStringName, string SqlText)
            {
                object Result = new object();
                string dbConnection = ConfigurationManager.ConnectionStrings[ConnectionStringName].ConnectionString;
                SqlConnection TempConnection = new SqlConnection(dbConnection);//连接字符串
                try
                {
                    SqlDataAdapter ToRun = new SqlDataAdapter();  //創建SqlDataAdapter 类

                    ToRun.SelectCommand = new SqlCommand(SqlText, TempConnection);
                    TempConnection.Open();
                    ToRun.SelectCommand.CommandType = System.Data.CommandType.Text;
                    Result = ToRun.SelectCommand.ExecuteScalar();
                }
                catch (Exception AnyError)
                {
                    throw AnyError;
                }
                finally
                {
                    TempConnection.Close();
                }


                return Result;
            }
            public static DataTable RunSqlDataTable(string ConnectionStringName, string SqlText)
            {

                DataTable Result = new DataTable();
                string dbConnection = ConfigurationManager.ConnectionStrings[ConnectionStringName].ConnectionString;
                SqlConnection TempConnection = new SqlConnection(dbConnection);//连接字符串
                try
                {
                    SqlDataAdapter ToRun = new SqlDataAdapter(SqlText, TempConnection);  //創建SqlDataAdapter 类
                    TempConnection.Open();
                    ToRun.Fill(Result);

                }
                catch (Exception AnyError)
                {
                    throw AnyError;
                }
                finally
                {
                    TempConnection.Close();
                }


                return Result;
            }
        }

    public class Util_File
    {
        public static string ReadToEnd(String FilePath,Encoding Enc)
        {
            FileStream fs = new FileStream(FilePath,FileMode.Open);
            byte[] buf = new byte[fs.Length];
            fs.Read(buf, 0, buf.Length);
            fs.Close();
            return Enc.GetString(buf);
        }
        public static void SaveToFile(String Content,String FilePath, Encoding Enc)
        {
            FileStream fs = null;
            if (File.Exists(FilePath))
            {
                fs = new FileStream(FilePath, FileMode.Truncate);
            }
            else
            {
                fs = new FileStream(FilePath, FileMode.Create);
            }
            byte[] bfs = Enc.GetBytes(Content);
            fs.Write(bfs,0,bfs.Length);
            fs.Flush();
            fs.Close();
        }
    }
    public class Util_WEB
    {
        public static string CurrentUrl = "";
        public static string OpenUrl(string TargetURL, string RefURL, string Body, string Method, System.Net.CookieCollection BrowCookie, bool AllowRedirect = true, bool KeepAlive = false, string ContentType = "application/json;charset=UTF-8", string authorization = "", Int32 TimeOut = 15000)
        {
            DateTime Pre = DateTime.Now;
            string Result = OpenUrl(TargetURL, RefURL, Body, Method, BrowCookie, Encoding.UTF8, AllowRedirect, KeepAlive, ContentType, authorization, TimeOut);
            NetFramework.Console.WriteLine("--------------网站下载总时间：" + (DateTime.Now - Pre).TotalSeconds.ToString(), false);
            return Result;
        }

        public static bool CheckValidationResult(object sender
            , System.Security.Cryptography.X509Certificates.X509Certificate certificate
            , System.Security.Cryptography.X509Certificates.X509Chain chain
            , System.Net.Security.SslPolicyErrors errors)
        {
            return true;
        }

        public static void SetHeaderValue(WebHeaderCollection header, string name, string value)
        {
            var property = typeof(WebHeaderCollection).GetProperty("InnerCollection",
                System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic);
            if (property != null)
            {
                var collection = property.GetValue(header, null) as System.Collections.Specialized.NameValueCollection;
                collection[name] = value;
            }
        }

        public static string OpenUrl(string TargetURL, string RefURL, string Body, string Method, System.Net.CookieCollection BrowCookie, Encoding ContactEncoding, bool AllowRedirect = false, bool KeepAlive = true, string ContentType = "application/json;charset=UTF-8", string authorization = "", Int32 TimeOut = 15000)
        {

            //System.Net.ServicePointManager.MaxServicePoints=20;

            System.Net.ServicePointManager.DefaultConnectionLimit = 500;

            System.Net.ServicePointManager.SetTcpKeepAlive(true, 15000, 15000);
            //HttpWebRequest LoginPage = null;
            //    GetHttpWebResponseNoRedirect(TargetURL,"","",out LoginPage);

            WebRequest LoginPage = HttpWebRequest.Create(TargetURL);
            ((HttpWebRequest)LoginPage).AllowAutoRedirect = AllowRedirect;
            //((HttpWebRequest)LoginPage).KeepAlive = KeepAlive;
            //SetHeaderValue(((HttpWebRequest)LoginPage).Headers, "Connection", "Keep-Alive");
            ((HttpWebRequest)LoginPage).Timeout = TimeOut;
            ((HttpWebRequest)LoginPage).Credentials = CredentialCache.DefaultCredentials;
            if (authorization != "")
            {
                LoginPage.Headers.Add("Authorization", authorization);

            }

            LoginPage.Method = Method;
            if (TargetURL.ToLower().StartsWith("https"))
            {
                //ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12 | SecurityProtocolType.Tls11 | SecurityProtocolType.Tls | SecurityProtocolType.Ssl3;

                //System.Net.ServicePointManager.ServerCertificateValidationCallback = CheckValidationResult;
                ((HttpWebRequest)LoginPage).ProtocolVersion = System.Net.HttpVersion.Version11;
            }

            switch (Method)
            {
                case "GET":
                    // ((HttpWebRequest)LoginPage).Accept = "application/x-ms-application, image/jpeg, application/xaml+xml, image/gif, image/pjpeg, application/x-ms-xbap, application/vnd.ms-powerpoint, application/msword, application/vnd.ms-excel,application/json, text/plain, */*";
                    ((HttpWebRequest)LoginPage).Accept = "*/*";
                    ((HttpWebRequest)LoginPage).UserAgent = "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/65.0.3325.181 Safari/537.36 OPR/52.0.2871.40";
                    LoginPage.Headers.Add("Accept-Encoding", "gzip, deflate,br");
                    LoginPage.Headers.Add("Accept-Language", "zh-CN,zh;q=0.9");
                    ((HttpWebRequest)LoginPage).CookieContainer = new CookieContainer();
                    ((HttpWebRequest)LoginPage).CookieContainer.Add(BrowCookie);

                    //((HttpWebRequest)LoginPage).Connection = "KeepAlive,Close";
                    ((HttpWebRequest)LoginPage).Referer = RefURL;

                    break;
                case "POST":
                    ((HttpWebRequest)LoginPage).Accept = "application/x-ms-application, image/jpeg, application/xaml+xml, image/gif, image/pjpeg, application/x-ms-xbap, application/vnd.ms-powerpoint, application/msword, application/vnd.ms-excel,application/json, text/plain, */*";
                    ((HttpWebRequest)LoginPage).Referer = RefURL;
                    ((HttpWebRequest)LoginPage).UserAgent = "Mozilla/4.0 (compatible; MSIE 8.0; Windows NT 6.1; WOW64; Trident/4.0; SLCC2; .NET CLR 2.0.50727; .NET CLR 3.5.30729; .NET CLR 3.0.30729; Media Center PC 6.0; .NET4.0C; .NET4.0E)";
                    LoginPage.Headers.Add("Accept-Encoding", "gzip, deflate,br");
                    LoginPage.Headers.Add("Accept-Language", "zh-CN,zh;q=0.9");
                    ((HttpWebRequest)LoginPage).CookieContainer = new CookieContainer();
                    ((HttpWebRequest)LoginPage).CookieContainer.Add(BrowCookie);
                    ((HttpWebRequest)LoginPage).ContentType = ContentType;
                    //((HttpWebRequest)LoginPage).ServicePoint.Expect100Continue = true;
                    //((HttpWebRequest)LoginPage).Connection = "KeepAlive";
                    if (((HttpWebRequest)LoginPage).Referer != null)
                    {
                        LoginPage.Headers.Add("Origin", ((HttpWebRequest)LoginPage).Referer.Substring(0, ((HttpWebRequest)LoginPage).Referer.Length - 1));

                    }

                    if (Body != "")
                    {
                        Stream bodys = LoginPage.GetRequestStream();

                        byte[] text = ContactEncoding.GetBytes(Body);

                        bodys.Write(text, 0, text.Length);

                        bodys.Flush();
                        bodys.Close();
                    }
                    break;
                case "OPTIONS":
                    // ((HttpWebRequest)LoginPage).Accept = "application/x-ms-application, image/jpeg, application/xaml+xml, image/gif, image/pjpeg, application/x-ms-xbap, application/vnd.ms-powerpoint, application/msword, application/vnd.ms-excel,application/json, text/plain, */*";
                    ((HttpWebRequest)LoginPage).Accept = "*/*";
                    ((HttpWebRequest)LoginPage).UserAgent = "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/65.0.3325.181 Safari/537.36 OPR/52.0.2871.40";
                    LoginPage.Headers.Add("Accept-Encoding", "gzip, deflate,br");
                    LoginPage.Headers.Add("Accept-Language", "zh-CN,zh;q=0.9");
                    ((HttpWebRequest)LoginPage).CookieContainer = new CookieContainer();
                    ((HttpWebRequest)LoginPage).CookieContainer.Add(BrowCookie);

                    //((HttpWebRequest)LoginPage).Connection = "KeepAlive";
                    ((HttpWebRequest)LoginPage).Referer = RefURL;
                    LoginPage.Headers.Add("Origin", RefURL);

                    break;
                default:
                    break;
            }
            //((HttpWebRequest)LoginPage).KeepAlive = true;
            SetHeaderValue(((HttpWebRequest)LoginPage).Headers, "Connection", "Keep-Alive");
            LoginPage.Timeout = TimeOut;
            if (RefURL.ToLower().StartsWith("https"))
            {
                //System.Net.ServicePointManager.ServerCertificateValidationCallback = CheckValidationResult;
                ((HttpWebRequest)LoginPage).ProtocolVersion = System.Net.HttpVersion.Version11;
            }
            //System.GC.Collect();
            System.Threading.Thread.Sleep(100);
            HttpWebResponse LoginPage_Return = null;
            try
            {
                CurrentUrl = "正在下载" + TargetURL;
                //System.GC.Collect();

                // NetFramework.Console.WriteLine("下载URL" + LoginPage.RequestUri.AbsoluteUri + Environment.NewLine);
                LoginPage_Return = (HttpWebResponse)LoginPage.GetResponse();

                CurrentUrl = "已下载" + TargetURL;

                if (((HttpWebResponse)LoginPage_Return).Headers["Set-Cookie"] != null)
                {
                    string Start = LoginPage.RequestUri.Host.Substring(0, LoginPage.RequestUri.Host.LastIndexOf("."));
                    string Host = LoginPage.RequestUri.Host.Substring(LoginPage.RequestUri.Host.LastIndexOf("."));

                    foreach (Cookie cookieitem in ((HttpWebResponse)LoginPage_Return).Cookies)
                    {
                        string[] SplitDomain = cookieitem.Domain.Split((".").ToCharArray(), StringSplitOptions.RemoveEmptyEntries);
                        Int32 Length = SplitDomain.Length;
                        cookieitem.Domain = "." + SplitDomain[Length - 2] + "." + SplitDomain[Length - 1];
                        cookieitem.Expires = cookieitem.Expires == null ? DateTime.Now.AddHours(168) : cookieitem.Expires.AddHours(168);
                        BrowCookie.Add(cookieitem);
                    }


                    //CookieContainer NC = new CookieContainer();
                    //NC.SetCookies(((HttpWebResponse)LoginPage_Return).ResponseUri, ((HttpWebResponse)LoginPage_Return).Headers["Set-Cookie"]);
                    //BrowCookie.Add(NC.GetCookies(((HttpWebResponse)LoginPage_Return).ResponseUri));

                    // Host = Start.Substring(Start.LastIndexOf(".")) + Host;
                    // AddCookieWithCookieHead(tmpcookie, ((HttpWebResponse)LoginPage_Return).Headers["Set-Cookie"].Replace("Secure,", ""), Host);
                }

            }
            catch (Exception AnyError)
            {

                LoginPage = null;
                System.GC.Collect();

                NetFramework.Console.WriteLine("网址打开失败" + TargetURL, true);
                NetFramework.Console.WriteLine("网址打开失败" + AnyError.Message, true);
                NetFramework.Console.WriteLine("网址打开失败" + AnyError.StackTrace, true);
                return "";
            }

            string responseBody = string.Empty;
            try
            {


                if (LoginPage_Return.ContentEncoding.ToLower().Contains("gzip"))
                {
                    using (GZipStream stream = new GZipStream(LoginPage_Return.GetResponseStream(), CompressionMode.Decompress))
                    {

                        using (StreamReader reader = new StreamReader(stream, ContactEncoding))
                        {
                            responseBody = reader.ReadToEnd();
                            stream.Close();
                        }
                    }
                }
                else if (LoginPage_Return.ContentEncoding.ToLower().Contains("deflate"))
                {
                    using (DeflateStream stream = new DeflateStream(LoginPage_Return.GetResponseStream(), CompressionMode.Decompress))
                    {
                        using (StreamReader reader = new StreamReader(stream, ContactEncoding))
                        {
                            responseBody = reader.ReadToEnd();
                            stream.Close();
                        }
                    }
                }
                else if (LoginPage_Return.ContentEncoding.ToLower().Contains("br"))
                {
                    using (Brotli.BrotliStream stream = new Brotli.BrotliStream(LoginPage_Return.GetResponseStream(), CompressionMode.Decompress))
                    {
                        using (StreamReader reader = new StreamReader(stream, ContactEncoding))
                        {
                            responseBody = reader.ReadToEnd();
                            stream.Close();
                        }
                    }
                }
                else
                {
                    using (Stream stream = LoginPage_Return.GetResponseStream())
                    {
                        using (StreamReader reader = new StreamReader(stream, ContactEncoding))
                        {
                            responseBody = reader.ReadToEnd();
                            stream.Close();
                        }
                    }
                }
            }
            catch (Exception AnyError)
            {
                LoginPage.Abort();
                LoginPage = null;
                //System.GC.Collect();

                NetFramework.Console.WriteLine("网址打开失败" + TargetURL, true);
                NetFramework.Console.WriteLine("网址打开失败" + AnyError.Message, true);
                NetFramework.Console.WriteLine("网址打开失败" + AnyError.StackTrace, true);
                return "";
            }
            LoginPage.Abort();


            //  NetFramework.Console.WriteLine("下载完成" + LoginPage_Return.ResponseUri.AbsoluteUri + Environment.NewLine);

            LoginPage_Return.Close();
            LoginPage_Return = null;
            LoginPage = null;
            System.GC.Collect();


            return responseBody;

        }

        private static DateTime? ImageToday = null;
        private static Int32 ImageFileid = 0;
        private static string Boundary = "wWMqeF7OGA3s1GXQ";

        public static string UploadWXImage(string ImgFilePath, string MyUserID, string TOUserID, string JavaTimeSpan, CookieCollection tmpcookie, Newtonsoft.Json.Linq.JObject RequestBase, string webhost)
        {
            #region 上传文件


            if (ImageToday == null || ImageToday != DateTime.Today || ImageFileid > 9)
            {
                ImageToday = DateTime.Today;
                ImageFileid = 1;
                Boundary = GenerateRandom(16);

            }

            FileInfo fi = new FileInfo(ImgFilePath);
            //POST /cgi-bin/mmwebwx-bin/webwxuploadmedia?f=json HTTP/1.1
            //Content-Type: multipart/form-data; boundary=----WebKitFormBoundarywWMqeF7OGA3s1GXQ

            //Host: file."+webhost+"





            string UploadUrl = "https://file." + webhost + "/cgi-bin/mmwebwx-bin/webwxuploadmedia?f=json";
            System.Net.ServicePointManager.DefaultConnectionLimit = 500;
            System.Net.ServicePointManager.SetTcpKeepAlive(true, 5000, 5000);


            ServicePointManager.SecurityProtocol = (SecurityProtocolType)192 | (SecurityProtocolType)768 | (SecurityProtocolType)3072;
            //ServicePointManager.SecurityProtocol = SecurityProtocolType.;

            string optionurl = "https://file." + webhost + "/cgi-bin/mmwebwx-bin/webwxuploadmedia?f=json";
            WebRequest OptionPage = HttpWebRequest.Create(optionurl);
            ((HttpWebRequest)OptionPage).Method = "OPTIONS";
            ((HttpWebRequest)OptionPage).Accept = "application/x-ms-application, image/jpeg, application/xaml+xml, image/gif, image/pjpeg, application/x-ms-xbap, application/vnd.ms-powerpoint, application/msword, application/vnd.ms-excel,application/json, text/plain, */*";
            ((HttpWebRequest)OptionPage).Referer = "https://" + webhost + "/";
            ((HttpWebRequest)OptionPage).UserAgent = "Mozilla/4.0 (compatible; MSIE 8.0; Windows NT 6.1; WOW64; Trident/4.0; SLCC2; .NET CLR 2.0.50727; .NET CLR 3.5.30729; .NET CLR 3.0.30729; Media Center PC 6.0; .NET4.0C; .NET4.0E)";
            OptionPage.Headers.Add("Accept-Encoding", "gzip, deflate,br");
            OptionPage.Headers.Add("Accept-Language", "zh-CN,zh;q=0.9");
            ((HttpWebRequest)OptionPage).CookieContainer = new CookieContainer();
            ((HttpWebRequest)OptionPage).CookieContainer.Add(tmpcookie);

            ((HttpWebRequest)OptionPage).Credentials = CredentialCache.DefaultCredentials;


            //((HttpWebRequest)OptionPage).Connection = "KeepAlive,Close";
            OptionPage.Headers.Add("Origin", ((HttpWebRequest)OptionPage).Referer.Substring(0, ((HttpWebRequest)OptionPage).Referer.Length - 1));

            StreamReader OptionReader = new StreamReader(OptionPage.GetResponse().GetResponseStream());
            string OptionResult = OptionReader.ReadToEnd();


            WebRequest LoginPage = HttpWebRequest.Create(UploadUrl);
            ((HttpWebRequest)LoginPage).Credentials = CredentialCache.DefaultCredentials;


            ((HttpWebRequest)LoginPage).AllowAutoRedirect = false;

            ((HttpWebRequest)LoginPage).Method = "POST";
            ((HttpWebRequest)LoginPage).Accept = "application/x-ms-application, image/jpeg, application/xaml+xml, image/gif, image/pjpeg, application/x-ms-xbap, application/vnd.ms-powerpoint, application/msword, application/vnd.ms-excel,application/json, text/plain, */*";
            ((HttpWebRequest)LoginPage).Referer = "https://" + webhost + "/";
            ((HttpWebRequest)LoginPage).UserAgent = "Mozilla/4.0 (compatible; MSIE 8.0; Windows NT 6.1; WOW64; Trident/4.0; SLCC2; .NET CLR 2.0.50727; .NET CLR 3.5.30729; .NET CLR 3.0.30729; Media Center PC 6.0; .NET4.0C; .NET4.0E)";
            LoginPage.Headers.Add("Accept-Encoding", "gzip, deflate,br");
            LoginPage.Headers.Add("Accept-Language", "zh-CN,zh;q=0.9");
            ((HttpWebRequest)LoginPage).CookieContainer = new CookieContainer();
            ((HttpWebRequest)LoginPage).CookieContainer.Add(tmpcookie);
            ((HttpWebRequest)LoginPage).ContentType = "multipart/form-data; boundary=----WebKitFormBoundary" + Boundary;

            // ((HttpWebRequest)LoginPage).Connection = "KeepAlive,Close";
            LoginPage.Headers.Add("Origin", ((HttpWebRequest)LoginPage).Referer.Substring(0, ((HttpWebRequest)LoginPage).Referer.Length - 1));
            Stream Strem_ToPost = LoginPage.GetRequestStream();

            //数据
            //------WebKitFormBoundarywWMqeF7OGA3s1GXQ
            //Content-Disposition: form-data; name="id"
            //WU_FILE_2
            byte[] buf = Encoding.UTF8.GetBytes("------WebKitFormBoundary" + Boundary + Environment.NewLine);
            Strem_ToPost.Write(buf, 0, buf.Length);
            buf = Encoding.UTF8.GetBytes("Content-Disposition: form-data; name=\"id\"" + Environment.NewLine + Environment.NewLine);
            Strem_ToPost.Write(buf, 0, buf.Length);

            buf = Encoding.UTF8.GetBytes("WU_FILE_" + ImageFileid.ToString() + Environment.NewLine);
            ImageFileid += 1;
            Strem_ToPost.Write(buf, 0, buf.Length);


            //------WebKitFormBoundarywWMqeF7OGA3s1GXQ
            //Content-Disposition: form-data; name="name"
            //Data.jpg
            buf = Encoding.UTF8.GetBytes("------WebKitFormBoundary" + Boundary + Environment.NewLine);
            Strem_ToPost.Write(buf, 0, buf.Length);
            buf = Encoding.UTF8.GetBytes("Content-Disposition: form-data; name=\"name\"" + Environment.NewLine + Environment.NewLine);
            Strem_ToPost.Write(buf, 0, buf.Length);
            buf = Encoding.UTF8.GetBytes(fi.Name + Environment.NewLine);
            Strem_ToPost.Write(buf, 0, buf.Length);
            //------WebKitFormBoundarywWMqeF7OGA3s1GXQ
            //Content-Disposition: form-data; name="type"
            //image/jpeg
            buf = Encoding.UTF8.GetBytes("------WebKitFormBoundary" + Boundary + Environment.NewLine);
            Strem_ToPost.Write(buf, 0, buf.Length);
            buf = Encoding.UTF8.GetBytes("Content-Disposition: form-data; name=\"type\"" + Environment.NewLine + Environment.NewLine);
            Strem_ToPost.Write(buf, 0, buf.Length);
            buf = Encoding.UTF8.GetBytes("image/jpeg" + Environment.NewLine);
            Strem_ToPost.Write(buf, 0, buf.Length);
            //------WebKitFormBoundarywWMqeF7OGA3s1GXQ
            //Content-Disposition: form-data; name="lastModifiedDate"
            //Mon Apr 09 2018 17:40:22 GMT+0800 (中国标准时间)
            buf = Encoding.UTF8.GetBytes("------WebKitFormBoundary" + Boundary + Environment.NewLine);
            Strem_ToPost.Write(buf, 0, buf.Length);
            buf = Encoding.UTF8.GetBytes("Content-Disposition: form-data; name=\"lastModifiedDate\"" + Environment.NewLine + Environment.NewLine);
            Strem_ToPost.Write(buf, 0, buf.Length);

            buf = Encoding.UTF8.GetBytes(fi.LastWriteTime.ToString("r") + Environment.NewLine);
            Strem_ToPost.Write(buf, 0, buf.Length);
            //------WebKitFormBoundarywWMqeF7OGA3s1GXQ
            //Content-Disposition: form-data; name="size"
            //79253
            buf = Encoding.UTF8.GetBytes("------WebKitFormBoundary" + Boundary + Environment.NewLine);
            Strem_ToPost.Write(buf, 0, buf.Length);
            buf = Encoding.UTF8.GetBytes("Content-Disposition: form-data; name=\"size\"" + Environment.NewLine + Environment.NewLine);
            Strem_ToPost.Write(buf, 0, buf.Length);
            buf = Encoding.UTF8.GetBytes(fi.Length.ToString() + Environment.NewLine);
            Strem_ToPost.Write(buf, 0, buf.Length);


            //------WebKitFormBoundarywWMqeF7OGA3s1GXQ
            //Content-Disposition: form-data; name="mediatype"
            //pic
            buf = Encoding.UTF8.GetBytes("------WebKitFormBoundary" + Boundary + Environment.NewLine);
            Strem_ToPost.Write(buf, 0, buf.Length);
            buf = Encoding.UTF8.GetBytes("Content-Disposition: form-data; name=\"mediatype\"" + Environment.NewLine + Environment.NewLine);
            Strem_ToPost.Write(buf, 0, buf.Length);
            buf = Encoding.UTF8.GetBytes("pic" + Environment.NewLine);
            Strem_ToPost.Write(buf, 0, buf.Length);
            //------WebKitFormBoundarywWMqeF7OGA3s1GXQ
            //Content-Disposition: form-data; name="uploadmediarequest"
            //{"UploadType":2,"BaseRequest":{"Uin":2402981522,"Sid":"W8Ia83fMnlcuKK0U","Skey":"@crypt_bbd454c7_9465a672aa848c64c765ea727877bdd1","DeviceID":"e718028710913369"},"ClientMediaId":1523267109886,"TotalLen":79253,"StartPos":0,"DataLen":79253,"MediaType":4,"FromUserName":"@ac0308d92ae0d88beb8d90feee45a86c02f36bd5f3560398b544abeac4e70a14","ToUserName":"@@f2a3e52ae022d3303864e9dfcda631635a429b2c28c4d93388ec25429280df00","FileMd5":"a5a03dda3342443cfd15c2ccc8f970e5"}
            buf = Encoding.UTF8.GetBytes("------WebKitFormBoundary" + Boundary + Environment.NewLine);
            Strem_ToPost.Write(buf, 0, buf.Length);
            buf = Encoding.UTF8.GetBytes("Content-Disposition: form-data; name=\"uploadmediarequest\"" + Environment.NewLine + Environment.NewLine);
            Strem_ToPost.Write(buf, 0, buf.Length);
            Newtonsoft.Json.Linq.JObject J_ToPost = new Newtonsoft.Json.Linq.JObject();
            J_ToPost.Add("UploadType", 2);
            J_ToPost.Add("BaseRequest", RequestBase["BaseRequest"]);

            J_ToPost.Add("ClientMediaId", JavaTimeSpan);
            J_ToPost.Add("TotalLen", fi.Length);
            J_ToPost.Add("StartPos", 0);
            J_ToPost.Add("DataLen", fi.Length);
            J_ToPost.Add("MediaType", 4);
            J_ToPost.Add("FromUserName", MyUserID);
            J_ToPost.Add("ToUserName", TOUserID);
            J_ToPost.Add("FileMd5", Util_MD5.GetMD5HashFromFile(ImgFilePath));



            buf = Encoding.UTF8.GetBytes(J_ToPost.ToString().Replace(Environment.NewLine, "") + Environment.NewLine);
            Strem_ToPost.Write(buf, 0, buf.Length);
            //------WebKitFormBoundarywWMqeF7OGA3s1GXQ
            //Content-Disposition: form-data; name="webwx_data_ticket"
            //gScs3xfj201uhj/fk3wxSMQA
            buf = Encoding.UTF8.GetBytes("------WebKitFormBoundary" + Boundary + Environment.NewLine);
            Strem_ToPost.Write(buf, 0, buf.Length);
            buf = Encoding.UTF8.GetBytes("Content-Disposition: form-data; name=\"webwx_data_ticket\"" + Environment.NewLine + Environment.NewLine);
            Strem_ToPost.Write(buf, 0, buf.Length);
            buf = Encoding.UTF8.GetBytes(tmpcookie["webwx_data_ticket"].Value + Environment.NewLine);
            Strem_ToPost.Write(buf, 0, buf.Length);

            //------WebKitFormBoundarywWMqeF7OGA3s1GXQ
            //Content-Disposition: form-data; name="pass_ticket"
            //undefined
            buf = Encoding.UTF8.GetBytes("------WebKitFormBoundary" + Boundary + Environment.NewLine);
            Strem_ToPost.Write(buf, 0, buf.Length);
            buf = Encoding.UTF8.GetBytes("Content-Disposition: form-data; name=\"pass_ticket\"" + Environment.NewLine + Environment.NewLine);
            Strem_ToPost.Write(buf, 0, buf.Length);
            buf = Encoding.UTF8.GetBytes(
                (tmpcookie["pass_ticket"] == null ? "undefined" : tmpcookie["pass_ticket"].Value) + Environment.NewLine
                );
            Strem_ToPost.Write(buf, 0, buf.Length);
            //------WebKitFormBoundarywWMqeF7OGA3s1GXQ
            //Content-Disposition: form-data; name="filename"; filename="Data.jpg"
            //Content-Type: image/jpeg
            buf = Encoding.UTF8.GetBytes("------WebKitFormBoundary" + Boundary + Environment.NewLine);
            Strem_ToPost.Write(buf, 0, buf.Length);
            buf = Encoding.UTF8.GetBytes("Content-Disposition: form-data; name=\"filename\" filename=\"" + fi.Name + "\"" + Environment.NewLine);
            Strem_ToPost.Write(buf, 0, buf.Length);
            buf = Encoding.UTF8.GetBytes("Content-Type: image/jpeg" + Environment.NewLine + Environment.NewLine);
            Strem_ToPost.Write(buf, 0, buf.Length);

            FileStream fs = fi.OpenRead();
            int Read = fs.ReadByte();
            while (Read != -1)
            {
                Strem_ToPost.WriteByte((byte)Read);
                Read = fs.ReadByte();
            }
            buf = Encoding.UTF8.GetBytes(Environment.NewLine + "------WebKitFormBoundary" + Boundary);
            Strem_ToPost.Write(buf, 0, buf.Length);

            buf = Encoding.UTF8.GetBytes("--" + Environment.NewLine);
            Strem_ToPost.Write(buf, 0, buf.Length);


            Strem_ToPost.Flush();
            Strem_ToPost.Close();

            //((HttpWebRequest)LoginPage).KeepAlive = true;
            //SetHeaderValue(((HttpWebRequest)LoginPage).Headers, "Connection", "Keep-Alive");
            ((HttpWebRequest)LoginPage).Timeout = 15000;

            System.GC.Collect();

            HttpWebResponse LoginPage_Return = null;
            try
            {
                NetFramework.Util_WEB.CurrentUrl = "正在下载" + UploadUrl;
                NetFramework.Console.WriteLine("正在上传图片", false);

                LoginPage_Return = (HttpWebResponse)LoginPage.GetResponse();
                NetFramework.Util_WEB.CurrentUrl = "已下载" + UploadUrl;
                NetFramework.Console.WriteLine("上传图片完成", false);
                if (((HttpWebResponse)LoginPage_Return).Headers["Set-Cookie"] != null)
                {
                    string Start = LoginPage.RequestUri.Host.Substring(0, LoginPage.RequestUri.Host.LastIndexOf("."));
                    string Host = LoginPage.RequestUri.Host.Substring(LoginPage.RequestUri.Host.LastIndexOf("."));

                    //tmpcookie.Add(((HttpWebResponse)LoginPage_Return).Cookies);

                    foreach (Cookie cookieitem in ((HttpWebResponse)LoginPage_Return).Cookies)
                    {
                        string[] SplitDomain = cookieitem.Domain.Split((".").ToCharArray(), StringSplitOptions.RemoveEmptyEntries);
                        Int32 Length = SplitDomain.Length;
                        cookieitem.Domain = "." + SplitDomain[Length - 2] + "." + SplitDomain[Length - 1];
                        tmpcookie.Add(cookieitem);
                    }

                    //CookieContainer NC = new CookieContainer();
                    //NC.SetCookies(((HttpWebResponse)LoginPage_Return).ResponseUri, ((HttpWebResponse)LoginPage_Return).Headers["Set-Cookie"]);
                    //  tmpcookie.Add(NC.GetCookies(((HttpWebResponse)LoginPage_Return).ResponseUri));

                    // Host = Start.Substring(Start.LastIndexOf(".")) + Host;
                    //AddCookieWithCookieHead(tmpcookie, ((HttpWebResponse)LoginPage_Return).Headers["Set-Cookie"].Replace("Secure,", ""), Host);
                }

            }
            catch (Exception AnyError)
            {

                throw AnyError;
            }

            string responseBody = string.Empty;
            if (LoginPage_Return.ContentEncoding.ToLower().Contains("gzip"))
            {
                using (System.IO.Compression.GZipStream stream = new System.IO.Compression.GZipStream(LoginPage_Return.GetResponseStream(), CompressionMode.Decompress))
                {
                    using (StreamReader reader = new StreamReader(stream))
                    {
                        responseBody = reader.ReadToEnd();
                        stream.Close();
                    }
                }
            }
            else if (LoginPage_Return.ContentEncoding.ToLower().Contains("deflate"))
            {
                using (System.IO.Compression.DeflateStream stream = new System.IO.Compression.DeflateStream(LoginPage_Return.GetResponseStream(), CompressionMode.Decompress))
                {
                    using (StreamReader reader = new StreamReader(stream, Encoding.UTF8))
                    {
                        responseBody = reader.ReadToEnd();
                        stream.Close();
                    }
                }
            }
            else
            {
                using (Stream stream = LoginPage_Return.GetResponseStream())
                {
                    using (StreamReader reader = new StreamReader(stream, Encoding.UTF8))
                    {
                        responseBody = reader.ReadToEnd();
                        stream.Close();
                    }
                }
            }
            LoginPage.Abort();

            LoginPage_Return.Close();
            LoginPage_Return = null;
            LoginPage = null;
            System.GC.Collect();
            NetFramework.Console.WriteLine("图片返回:" + responseBody, false);
            return responseBody;
            //返回：
            //            {
            //"BaseResponse": {
            //"Ret": 0,
            //"ErrMsg": ""
            //}
            //,
            //"MediaId": "@crypt_d4a4de27_1bf919affd8946a58cfbaad9eeea3cc0eb86c6890d6d66227dff9d3ace4cb1b4ec8db12709140e6c0bfdfc6b92c27e3c4426225ddc43c9241aacaf18dc8a5f92bc13106caccea22ba76b324c9a796ef1377e83a585b8ceab687df00db68f39fee1d7531b1594737c44e379a4b4d539c466d377a749f21ae1dd3917dcaca2c5ba223d3034eb193a4258dca898ad4aa9d5b1e356eb5e7879bea7b9a0897f1f7f96a6acdb2ea255f9a8873cacc1c3fa827ca5a7c9182e149ca80c5ff2d2a0048fdb8a1c0e61b6a3cc3eb5902f7a4f9b524983eefc37bc84e69f5374898f3312d615022d188fd04b91b0e6be51118a3e7df645512d6c5419e80f32584a3bbec8692179478ea4ee6c4a85c99ec92d8d1ba965a1be94aaf3fb5f5de702bf519aacc073242189f72616b7590fe94986b43a395b63a8d889b6e82375d472cc57df1c0422",
            //"StartPos": 79253,
            //"CDNThumbImgHeight": 100,
            //"CDNThumbImgWidth": 74,
            //"EncryFileName": "Data%2Ejpg"
            //}


            #endregion





            //POST /cgi-bin/mmwebwx-bin/webwxsendmsgimg?fun=async&f=json HTTP/1.1
            //Host: "+webhost+"
            //        {
            //    "BaseRequest": {
            //        "Uin": 2402981522,
            //        "Sid": "W8Ia83fMnlcuKK0U",
            //        "Skey": "@crypt_bbd454c7_9465a672aa848c64c765ea727877bdd1",
            //        "DeviceID": "e871841233370548"
            //    },
            //    "Msg": {
            //        "Type": 3,
            //        "MediaId": "@crypt_d4a4de27_1bf919affd8946a58cfbaad9eeea3cc0eb86c6890d6d66227dff9d3ace4cb1b4ec8db12709140e6c0bfdfc6b92c27e3c4426225ddc43c9241aacaf18dc8a5f92bc13106caccea22ba76b324c9a796ef1377e83a585b8ceab687df00db68f39fee1d7531b1594737c44e379a4b4d539c466d377a749f21ae1dd3917dcaca2c5ba223d3034eb193a4258dca898ad4aa9d5b1e356eb5e7879bea7b9a0897f1f7f96a6acdb2ea255f9a8873cacc1c3fa827ca5a7c9182e149ca80c5ff2d2a0048fdb8a1c0e61b6a3cc3eb5902f7a4f9b524983eefc37bc84e69f5374898f3312d615022d188fd04b91b0e6be51118a3e7df645512d6c5419e80f32584a3bbec8692179478ea4ee6c4a85c99ec92d8d1ba965a1be94aaf3fb5f5de702bf519aacc073242189f72616b7590fe94986b43a395b63a8d889b6e82375d472cc57df1c0422",
            //        "Content": "",
            //        "FromUserName": "@ac0308d92ae0d88beb8d90feee45a86c02f36bd5f3560398b544abeac4e70a14",
            //        "ToUserName": "@@f2a3e52ae022d3303864e9dfcda631635a429b2c28c4d93388ec25429280df00",
            //        "LocalID": "15232671098860552",
            //        "ClientMsgId": "15232671098860552"
            //    },
            //    "Scene": 0
            //}


        }


        public static string UploadYixinImage(string ImgFilePath, CookieCollection tmpcookie, string UserID, string Sessionid)
        {
            #region 上传文件



            FileInfo fi = new FileInfo(ImgFilePath);
            //POST /cgi-bin/mmwebwx-bin/webwxuploadmedia?f=json HTTP/1.1
            //Content-Type: multipart/form-data; boundary=----WebKitFormBoundarywWMqeF7OGA3s1GXQ

            //Host: file."+webhost+"

            string optionurl = "https://nos-hz.yixin.im/nos/webbatchupload?uid=" + UserID + "&sid=" + Sessionid + "&size=1&type=0&limit=15";
            WebRequest OptionPage = HttpWebRequest.Create(optionurl);
            ((HttpWebRequest)OptionPage).Method = "OPTIONS";
            ((HttpWebRequest)OptionPage).Accept = "application/x-ms-application, image/jpeg, application/xaml+xml, image/gif, image/pjpeg, application/x-ms-xbap, application/vnd.ms-powerpoint, application/msword, application/vnd.ms-excel,application/json, text/plain, */*";
            ((HttpWebRequest)OptionPage).Referer = "https://" + "web.yixin.im" + "/";
            ((HttpWebRequest)OptionPage).UserAgent = "Mozilla/4.0 (compatible; MSIE 8.0; Windows NT 6.1; WOW64; Trident/4.0; SLCC2; .NET CLR 2.0.50727; .NET CLR 3.5.30729; .NET CLR 3.0.30729; Media Center PC 6.0; .NET4.0C; .NET4.0E)";
            OptionPage.Headers.Add("Accept-Encoding", "gzip, deflate,br");
            OptionPage.Headers.Add("Accept-Language", "zh-CN,zh;q=0.9");
            ((HttpWebRequest)OptionPage).CookieContainer = new CookieContainer();
            ((HttpWebRequest)OptionPage).CookieContainer.Add(tmpcookie);

            // ((HttpWebRequest)OptionPage).Connection = "KeepAlive,Close";
            OptionPage.Headers.Add("Origin", ((HttpWebRequest)OptionPage).Referer.Substring(0, ((HttpWebRequest)OptionPage).Referer.Length - 1));

            StreamReader OptionReader = new StreamReader(OptionPage.GetResponse().GetResponseStream());
            string OptionResult = OptionReader.ReadToEnd();



            string UploadUrl = "https://nos-hz.yixin.im/nos/webbatchupload?uid=" + UserID + "&sid=" + Sessionid + "&size=1&type=0&limit=15";
            System.Net.ServicePointManager.DefaultConnectionLimit = 500;
            System.Net.ServicePointManager.SetTcpKeepAlive(true, 15000, 15000);


            ServicePointManager.SecurityProtocol = (SecurityProtocolType)192 | (SecurityProtocolType)768 | (SecurityProtocolType)3072;
            //ServicePointManager.SecurityProtocol = SecurityProtocolType.Ssl3;



            WebRequest LoginPage = HttpWebRequest.Create(UploadUrl);
            ((HttpWebRequest)LoginPage).AllowAutoRedirect = false;

            ((HttpWebRequest)LoginPage).Method = "POST";
            ((HttpWebRequest)LoginPage).Accept = "application/x-ms-application, image/jpeg, application/xaml+xml, image/gif, image/pjpeg, application/x-ms-xbap, application/vnd.ms-powerpoint, application/msword, application/vnd.ms-excel,application/json, text/plain, */*";
            ((HttpWebRequest)LoginPage).Referer = "https://" + "web.yixin.im" + "/";
            ((HttpWebRequest)LoginPage).UserAgent = "Mozilla/4.0 (compatible; MSIE 8.0; Windows NT 6.1; WOW64; Trident/4.0; SLCC2; .NET CLR 2.0.50727; .NET CLR 3.5.30729; .NET CLR 3.0.30729; Media Center PC 6.0; .NET4.0C; .NET4.0E)";
            LoginPage.Headers.Add("Accept-Encoding", "gzip, deflate,br");
            OptionPage.Headers.Add("Accept-Language", "zh-CN,zh;q=0.9");
            ((HttpWebRequest)LoginPage).CookieContainer = new CookieContainer();
            ((HttpWebRequest)LoginPage).CookieContainer.Add(tmpcookie);
            ((HttpWebRequest)LoginPage).ContentType = "multipart/form-data; boundary=----WebKitFormBoundary" + Boundary;

            // ((HttpWebRequest)LoginPage).Connection = "KeepAlive,Close";
            LoginPage.Headers.Add("Origin", "https://web.yixin.im");
            Stream Strem_ToPost = LoginPage.GetRequestStream();

            //数据

            //------WebKitFormBoundarywWMqeF7OGA3s1GXQ
            //Content-Disposition: form-data; name="filename"; filename="Data.jpg"
            //Content-Type: image/jpeg
            byte[] buf = Encoding.UTF8.GetBytes("------WebKitFormBoundary" + Boundary + Environment.NewLine);
            Strem_ToPost.Write(buf, 0, buf.Length);
            buf = Encoding.UTF8.GetBytes("Content-Disposition: form-data; name=\"" + fi.Name + "\"; filename=\"" + fi.Name + "\"" + Environment.NewLine);
            Strem_ToPost.Write(buf, 0, buf.Length);
            buf = Encoding.UTF8.GetBytes("Content-Type: image/jpeg" + Environment.NewLine + Environment.NewLine);
            Strem_ToPost.Write(buf, 0, buf.Length);

            FileStream fs = fi.OpenRead();
            int Read = fs.ReadByte();
            while (Read != -1)
            {
                Strem_ToPost.WriteByte((byte)Read);
                Read = fs.ReadByte();
            }
            buf = Encoding.UTF8.GetBytes(Environment.NewLine + "------WebKitFormBoundary" + Boundary);
            Strem_ToPost.Write(buf, 0, buf.Length);

            buf = Encoding.UTF8.GetBytes("--" + Environment.NewLine);
            Strem_ToPost.Write(buf, 0, buf.Length);


            Strem_ToPost.Flush();
            Strem_ToPost.Close();

            //((HttpWebRequest)LoginPage).KeepAlive = true;
            //SetHeaderValue(((HttpWebRequest)LoginPage).Headers, "Connection", "Keep-Alive");
            ((HttpWebRequest)LoginPage).Timeout = 15000;

            System.GC.Collect();

            HttpWebResponse LoginPage_Return = null;
            try
            {
                NetFramework.Util_WEB.CurrentUrl = "正在下载" + UploadUrl;
                NetFramework.Console.WriteLine("正在上传图片", false);

                LoginPage_Return = (HttpWebResponse)LoginPage.GetResponse();
                NetFramework.Util_WEB.CurrentUrl = "已下载" + UploadUrl;
                NetFramework.Console.WriteLine("上传图片完成", false);
                if (((HttpWebResponse)LoginPage_Return).Headers["Set-Cookie"] != null)
                {
                    string Start = LoginPage.RequestUri.Host.Substring(0, LoginPage.RequestUri.Host.LastIndexOf("."));
                    string Host = LoginPage.RequestUri.Host.Substring(LoginPage.RequestUri.Host.LastIndexOf("."));

                    // tmpcookie.Add(((HttpWebResponse)LoginPage_Return).Cookies);


                    foreach (Cookie cookieitem in ((HttpWebResponse)LoginPage_Return).Cookies)
                    {
                        string[] SplitDomain = cookieitem.Domain.Split((".").ToCharArray(), StringSplitOptions.RemoveEmptyEntries);
                        Int32 Length = SplitDomain.Length;
                        cookieitem.Domain = "." + SplitDomain[Length - 2] + "." + SplitDomain[Length - 1];
                        tmpcookie.Add(cookieitem);
                    }

                    //CookieContainer NC = new CookieContainer();
                    //NC.SetCookies(((HttpWebResponse)LoginPage_Return).ResponseUri, ((HttpWebResponse)LoginPage_Return).Headers["Set-Cookie"]);
                    // tmpcookie.Add(NC.GetCookies(((HttpWebResponse)LoginPage_Return).ResponseUri));


                    // Host = Start.Substring(Start.LastIndexOf(".")) + Host;
                    //AddCookieWithCookieHead(tmpcookie, ((HttpWebResponse)LoginPage_Return).Headers["Set-Cookie"].Replace("Secure,", ""), Host);
                }

            }
            catch (Exception AnyError)
            {

                throw AnyError;
            }

            string responseBody = string.Empty;
            if (LoginPage_Return.ContentEncoding.ToLower().Contains("gzip"))
            {
                using (System.IO.Compression.GZipStream stream = new System.IO.Compression.GZipStream(LoginPage_Return.GetResponseStream(), CompressionMode.Decompress))
                {
                    using (StreamReader reader = new StreamReader(stream))
                    {
                        responseBody = reader.ReadToEnd();
                        stream.Close();
                    }
                }
            }
            else if (LoginPage_Return.ContentEncoding.ToLower().Contains("deflate"))
            {
                using (System.IO.Compression.DeflateStream stream = new System.IO.Compression.DeflateStream(LoginPage_Return.GetResponseStream(), CompressionMode.Decompress))
                {
                    using (StreamReader reader = new StreamReader(stream, Encoding.UTF8))
                    {
                        responseBody = reader.ReadToEnd();
                        stream.Close();
                    }
                }
            }
            else
            {
                using (Stream stream = LoginPage_Return.GetResponseStream())
                {
                    using (StreamReader reader = new StreamReader(stream, Encoding.UTF8))
                    {
                        responseBody = reader.ReadToEnd();
                        stream.Close();
                    }
                }
            }
            LoginPage.Abort();

            LoginPage_Return.Close();
            LoginPage_Return = null;
            LoginPage = null;
            System.GC.Collect();
            NetFramework.Console.WriteLine("图片返回:" + responseBody, false);
            return responseBody;
            //返回：
            //            {
            //"BaseResponse": {
            //"Ret": 0,
            //"ErrMsg": ""
            //}
            //,
            //"MediaId": "@crypt_d4a4de27_1bf919affd8946a58cfbaad9eeea3cc0eb86c6890d6d66227dff9d3ace4cb1b4ec8db12709140e6c0bfdfc6b92c27e3c4426225ddc43c9241aacaf18dc8a5f92bc13106caccea22ba76b324c9a796ef1377e83a585b8ceab687df00db68f39fee1d7531b1594737c44e379a4b4d539c466d377a749f21ae1dd3917dcaca2c5ba223d3034eb193a4258dca898ad4aa9d5b1e356eb5e7879bea7b9a0897f1f7f96a6acdb2ea255f9a8873cacc1c3fa827ca5a7c9182e149ca80c5ff2d2a0048fdb8a1c0e61b6a3cc3eb5902f7a4f9b524983eefc37bc84e69f5374898f3312d615022d188fd04b91b0e6be51118a3e7df645512d6c5419e80f32584a3bbec8692179478ea4ee6c4a85c99ec92d8d1ba965a1be94aaf3fb5f5de702bf519aacc073242189f72616b7590fe94986b43a395b63a8d889b6e82375d472cc57df1c0422",
            //"StartPos": 79253,
            //"CDNThumbImgHeight": 100,
            //"CDNThumbImgWidth": 74,
            //"EncryFileName": "Data%2Ejpg"
            //}


            #endregion





            //POST /cgi-bin/mmwebwx-bin/webwxsendmsgimg?fun=async&f=json HTTP/1.1
            //Host: "+webhost+"
            //        {
            //    "BaseRequest": {
            //        "Uin": 2402981522,
            //        "Sid": "W8Ia83fMnlcuKK0U",
            //        "Skey": "@crypt_bbd454c7_9465a672aa848c64c765ea727877bdd1",
            //        "DeviceID": "e871841233370548"
            //    },
            //    "Msg": {
            //        "Type": 3,
            //        "MediaId": "@crypt_d4a4de27_1bf919affd8946a58cfbaad9eeea3cc0eb86c6890d6d66227dff9d3ace4cb1b4ec8db12709140e6c0bfdfc6b92c27e3c4426225ddc43c9241aacaf18dc8a5f92bc13106caccea22ba76b324c9a796ef1377e83a585b8ceab687df00db68f39fee1d7531b1594737c44e379a4b4d539c466d377a749f21ae1dd3917dcaca2c5ba223d3034eb193a4258dca898ad4aa9d5b1e356eb5e7879bea7b9a0897f1f7f96a6acdb2ea255f9a8873cacc1c3fa827ca5a7c9182e149ca80c5ff2d2a0048fdb8a1c0e61b6a3cc3eb5902f7a4f9b524983eefc37bc84e69f5374898f3312d615022d188fd04b91b0e6be51118a3e7df645512d6c5419e80f32584a3bbec8692179478ea4ee6c4a85c99ec92d8d1ba965a1be94aaf3fb5f5de702bf519aacc073242189f72616b7590fe94986b43a395b63a8d889b6e82375d472cc57df1c0422",
            //        "Content": "",
            //        "FromUserName": "@ac0308d92ae0d88beb8d90feee45a86c02f36bd5f3560398b544abeac4e70a14",
            //        "ToUserName": "@@f2a3e52ae022d3303864e9dfcda631635a429b2c28c4d93388ec25429280df00",
            //        "LocalID": "15232671098860552",
            //        "ClientMsgId": "15232671098860552"
            //    },
            //    "Scene": 0
            //}

        }


        #region 从包含多个 Cookie 的字符串读取到 CookieCollection 集合中
        public static void AddCookieWithCookieHead(CookieCollection cookieCol, string cookieHead, string defaultDomain)
        {
            if (cookieHead == null) return;

            string[] ary = cookieHead.Split(new string[] { "GMT" }, StringSplitOptions.RemoveEmptyEntries);
            for (int i = 0; i < ary.Length; i++)
            {
                string CookieStr = ary[i].Trim() + "GMT";
                if (CookieStr.StartsWith(","))
                {
                    CookieStr = CookieStr.Substring(1);
                }
                Cookie ck = GetCookieFromString(CookieStr, defaultDomain);
                if (ck != null)
                {
                    cookieCol.Add(ck);
                }
            }


        }
        #endregion
        #region 读取某一个 Cookie 字符串到 Cookie 变量中
        private static Cookie GetCookieFromString(string cookieString, string defaultDomain)
        {
            string[] ary = cookieString.Split(";".ToCharArray(), StringSplitOptions.RemoveEmptyEntries);
            System.Collections.Hashtable hs = new System.Collections.Hashtable();
            for (int i = 0; i < ary.Length; i++)
            {
                string s = ary[i].Trim();
                int index = s.IndexOf("=");
                if (index > 0)
                {
                    hs.Add(s.Substring(0, index), s.Substring(index + 1));
                }
            }
            Cookie ck = new Cookie();
            foreach (object Key in hs.Keys)
            {
                if (Key.ToString().ToLower() == "path") ck.Path = hs[Key].ToString();

                else if (Key.ToString().ToLower() == "expires")
                {
                    //ck.Expires = DateTime.Parse(hs[Key].ToString());
                }
                else if (Key.ToString().ToLower() == "domain") ck.Domain = defaultDomain;//hs[Key].ToString();
                else
                {
                    ck.Name = Key.ToString();
                    ck.Value = hs[Key].ToString();
                }
            }
            if (ck.Name == "") return null;
            if (ck.Domain == "") ck.Domain = defaultDomain;
            return ck;
        }
        #endregion


        public static string CleanHtml(string strHtml)
        {
            if (string.IsNullOrEmpty(strHtml)) return strHtml;
            //删除脚本
            //Regex.Replace(strHtml, @"<script[^>]*?>.*?</script>", "", RegexOptions.IgnoreCase)
            strHtml = Regex.Replace(strHtml, "(<script(.+?)</script>)|(<style(.+?)</style>)", "", RegexOptions.IgnoreCase | RegexOptions.Singleline);
            //删除标签
            var r = new Regex(@"</?[^>]*>", RegexOptions.IgnoreCase);
            Match m;
            for (m = r.Match(strHtml); m.Success; m = m.NextMatch())
            {
                strHtml = strHtml.Replace(m.Groups[0].ToString(), "");
            }
            return strHtml.Trim().Replace("&nbsp;", "");
        }

        private static char[] constant =
{
'a','b','c','d','e','f','g','h','i','j','k','l','m','n','o','p','q','r','s','t','u','v','w','x','y','z',
'A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z'
};
        public static string GenerateRandom(int Length)
        {
            System.Text.StringBuilder newRandom = new System.Text.StringBuilder(52);
            Random rd = new Random();
            for (int i = 0; i < Length; i++)
            {
                newRandom.Append(constant[rd.Next(52)]);
            }
            return newRandom.ToString();
        }
    }

    public class Console
    {
        public static void Write(string Message)
        {
            //LastLog = Message + LastLog;
            //if (LastLog.Length > 5000)
            //{
            //    LastLog = LastLog.Substring(0, 5000);
            //}
            System.Console.WriteLine(Message);
        }
        public static void WriteLine(string Message, bool Exception)
        {
            //LastLog = Message + Environment.NewLine + LastLog;
            //if (LastLog.Length > 5000)
            //{
            //    LastLog = LastLog.Substring(0, 5000);
            //}
            if (Exception)
            {
                System.Console.WriteLine(Message);
            }


        }
        public static string LastLog = "";

    }
    public class Util_MD5
    {
        public static string GetMD5HashFromFile(string fileName)
        {
            try
            {
                FileStream file = new FileStream(fileName, System.IO.FileMode.OpenOrCreate);
                System.Security.Cryptography.MD5 md5 = new System.Security.Cryptography.MD5CryptoServiceProvider();
                byte[] retVal = md5.ComputeHash(file);
                file.Close();
                StringBuilder sb = new StringBuilder();
                for (int i = 0; i < retVal.Length; i++)
                {
                    sb.Append(retVal[i].ToString("x2"));
                }
                return sb.ToString();
            }
            catch (Exception ex)
            {
                throw new Exception("GetMD5HashFromFile() fail,error:" + ex.Message);
            }
        }
        public static string GetStrMd5X2(string ConvertString)
        {
            System.Security.Cryptography.MD5CryptoServiceProvider md5 = new System.Security.Cryptography.MD5CryptoServiceProvider();
            string t2 = BitConverter.ToString(md5.ComputeHash(UTF8Encoding.Default.GetBytes(ConvertString)));
            t2 = t2.Replace("-", "");
            return t2;
        }
        public static bool MD5Success(string Total, out DateTime? Value, Guid MyGuid)
        {
            byte[] Totalb = null;
            try
            {

                Totalb = Convert.FromBase64String(Total);



                string Text = Encoding.UTF8.GetString(Totalb);
                string Time = Text.Substring(Text.Length - 59);
                //yyyy-MM-dd HH:mm：ss FFFF
                //bf697c61-e1ef-4848-9f03-558ab55686e9 36位
                string MD5 = Text.Substring(0, Text.Length - 59);
                string CheckMD5 = GetStrMd5X2(Time);

                Guid Passid = new Guid(Time.Substring(0, 36));
                string OutTime = Time.Substring(36);
                if (CheckMD5 == MD5 && Passid == MyGuid)
                {
                    Value = DateTime.Parse(OutTime);
                    return true;
                }
                else
                {
                    Value = null;
                    return false;
                }
            }
            catch (Exception)
            {
                Value = null;
                return false;
            }
        }


    }

    public class Util_Math
    {
        public static bool IsNumber(String strNumber)
        {
            Regex objNotNumberPattern = new Regex("[^0-9.-]");
            Regex objTwoDotPattern = new Regex("[0-9]*[.][0-9]*[.][0-9]*");
            Regex objTwoMinusPattern = new Regex("[0-9]*[-][0-9]*[-][0-9]*");
            String strValidRealPattern = "^([-]|[.]|[-.]|[0-9])[0-9]*[.]*[0-9]+$";
            String strValidIntegerPattern = "^([-]|[0-9])[0-9]*$";
            Regex objNumberPattern = new Regex("(" + strValidRealPattern + ")|(" + strValidIntegerPattern + ")");
            return !objNotNumberPattern.IsMatch(strNumber) &&
                !objTwoDotPattern.IsMatch(strNumber) &&
                !objTwoMinusPattern.IsMatch(strNumber) &&
                objNumberPattern.IsMatch(strNumber);
        }

        public static decimal NullToZero(decimal? dbvalue, Int32 KeepCount = 0)
        {
            return dbvalue.HasValue ? Math.Round(dbvalue.Value, KeepCount) : 0;
        }
    }

    public class Util_Cofig
    {



        //加密web.Config中的指定节
        public static void ProtectSection(string sectionName, string Path)
        {
            System.Configuration.Configuration config = WebConfigurationManager.OpenWebConfiguration(Path);
            System.Configuration.ConfigurationSection section = config.GetSection(sectionName);
            if (section != null && !section.SectionInformation.IsProtected)
            {
                section.SectionInformation.ProtectSection("DataProtectionConfigurationProvider");
                config.Save();
            }
        }

        //解密web.Config中的指定节
        public static void UnProtectSection(string sectionName, string Path)
        {
            System.Configuration.Configuration config = WebConfigurationManager.OpenWebConfiguration(Path);
            System.Configuration.ConfigurationSection section = config.GetSection(sectionName);
            if (section != null && section.SectionInformation.IsProtected)
            {
                section.SectionInformation.UnprotectSection();
                config.Save();
            }
        }
    }

    public class Util_DataTable
    {
        /// <summary>
        /// DataTable序列化
        /// </summary>
        /// <param name="dt"></param>
        /// <returns></returns>
        public static string SerializeDataTable(DataTable dt)
        {
            JavaScriptSerializer serializer = new JavaScriptSerializer();
            List<Dictionary<string, object>> list = new List<Dictionary<string, object>>();
            foreach (DataRow dr in dt.Rows)//每一行信息，新建一个Dictionary<string,object>,将该行的每列信息加入到字典
            {
                Dictionary<string, object> result = new Dictionary<string, object>();
                foreach (DataColumn dc in dt.Columns)
                {
                    result.Add(dc.ColumnName, dr[dc].ToString());
                }
                list.Add(result);
            }
            return serializer.Serialize(list);//调用Serializer方法
        }
        /// <summary>
        /// DataTable反序列化
        /// </summary>
        /// <param name="strJson"></param>
        /// <returns></returns>
        public static DataTable DeserializerTable(string strJson)
        {
            DataTable dt = new DataTable();
            JavaScriptSerializer serializer = new JavaScriptSerializer();
            ArrayList arralList = serializer.Deserialize<ArrayList>(strJson);//反序列化ArrayList类型
            if (arralList.Count > 0)//反序列化后ArrayList个数不为0
            {
                foreach (Dictionary<string, object> row in arralList)
                {
                    if (dt.Columns.Count == 0)//新建的DataTable中无任何信息，为其添加列名及类型
                    {
                        foreach (string key in row.Keys)
                        {
                            dt.Columns.Add(key, row[key].GetType());//添加dt的列名
                        }
                    }
                    DataRow dr = dt.NewRow();
                    foreach (string key in row.Keys)//讲arrayList中的值添加到DataTable中
                    {

                        dr[key] = row[key];//添加列值
                    }
                    dt.Rows.Add(dr);//添加一行
                }
            }
            return dt;
        }
    }

}

