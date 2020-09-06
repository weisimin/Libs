using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net;
using System.Windows.Forms;
using System.Xml;

using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.HtmlControls;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;

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

/*XLS數據源引用*/
using System.Data.OleDb;

using System.Xml.Serialization;
using System.Diagnostics;
/*NPOI第三方引用*/
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.HSSF.UserModel;
/*添加集合引用*/
using System.Collections.Specialized;
/*生成缩略图引用*/
using System.Drawing;
using System.Drawing.Imaging;
using System.Drawing.Drawing2D;

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

    #region  Quoted-Printable 解码
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
    #region "XLS文件工具"
    public class Util_XLS
    {

        /// <summary>
        /// 執行查詢
        /// </summary>
        /// <param name="ServerFileName"></param>
        /// <param name="SelectSQL"></param>
        /// <returns></returns>
        public static DataSet SelectFromXLS(string ServerFileName, string SelectSQL)
        {
            string mystring = "Provider = Microsoft.Jet.OLEDB.4.0 ; Data Source = '" + ServerFileName + "';Extended Properties=\"Excel 8.0;HDR=Yes;IMEX=1\"";
            OleDbConnection cnnxls = new OleDbConnection(mystring);
            OleDbDataAdapter myDa = null;
            DataSet myDs = new DataSet();
            try
            {
                cnnxls.Open();
                myDa = new OleDbDataAdapter(SelectSQL, cnnxls);
                myDa.Fill(myDs, "SelectResult");
            }
            catch (Exception AnyError)
            {
                cnnxls.Close();
                try
                {
                    return SelectFromXLS2007_ODBC(ServerFileName, SelectSQL);
                }
                catch (Exception Error2)
                {

                    throw new Exception(SelectSQL + Error2.Message);
                }
                throw AnyError;
            }
            finally
            {
                cnnxls.Close();
            }
            return myDs;

        }

        public static DataSet SelectFromXLS2007_ODBC(string ServerFileName, string SelectSQL)
        {
            string mystring = "Provider = Microsoft.ACE.OLEDB.12.0; Data Source = '" + ServerFileName + "';Extended Properties='Excel 12.0 Xml;HDR=Yes;IMEX=1'";
            OleDbConnection cnnxls = new OleDbConnection(mystring);
            OleDbDataAdapter myDa = null;
            DataSet myDs = new DataSet();
            try
            {
                cnnxls.Open();
                myDa = new OleDbDataAdapter(SelectSQL, cnnxls);
                myDa.Fill(myDs, "SelectResult");


            }
            catch (Exception AnyError)
            {
                cnnxls.Close();

                throw AnyError;


            }
            finally
            {
                cnnxls.Close();
            }
            return myDs;

        }


        public static DataSet SelectFromExcelReader(string ServerFileName, string SelectSQL)
        {


            FileStream stream = File.Open(ServerFileName, FileMode.Open, FileAccess.Read);
            //Reading from a binary Excel file ('97-2003 format; *.xls)
            //新建IExcelDataReader对象
            ExcelDataReader.IExcelDataReader excelReader = null;
            if (ServerFileName.EndsWith(".xls"))
            {
                excelReader = ExcelDataReader.ExcelReaderFactory.CreateBinaryReader(stream, null);
            }
            else if (ServerFileName.EndsWith(".xlsx"))
            {
                excelReader = ExcelDataReader.ExcelReaderFactory.CreateOpenXmlReader(stream, null);
            }
            else if (ServerFileName.EndsWith(".csv"))
            {
                excelReader = ExcelDataReader.ExcelReaderFactory.CreateCsvReader(stream,
                     new ExcelDataReader.ExcelReaderConfiguration()
                     {
                         FallbackEncoding = System.Text.Encoding.Default
                     }
                    );
            }

            //取得excel文档中所有的Sheet表
            DataSet result =
                ExcelDataReader.ExcelDataReaderExtensions.AsDataSet(
            excelReader, new ExcelDataReader.ExcelDataSetConfiguration()
            {
                ConfigureDataTable = (tableReader) => new ExcelDataReader.ExcelDataTableConfiguration()
                {
                    UseHeaderRow = true
                }
            });

            //关闭IExcelDataReader对象
            excelReader.Close();
            //将第一个sheet表赋值给dataGrid，在dataGrid中显示

            #region 转化SQL表格为SelectResult ,以及as 字段处理
            string dolartable = SelectSQL.Substring(0, SelectSQL.LastIndexOf("$"));
            string table = dolartable.Substring(dolartable.LastIndexOf("[") + 1);
            if (ServerFileName.EndsWith(".csv"))
            {
                result.Tables[0].TableName = table;
            }
            foreach (DataTable tableitem in result.Tables)
            {
                if (tableitem.TableName == table)
                {
                    tableitem.TableName = "SelectResult";

                    string Cols = dolartable.Substring(0, dolartable.LastIndexOf("from") - 1).Trim().Replace("\t", "");
                    string[] DoCols = Cols.Split(",".ToCharArray());
                    foreach (var Colitem in DoCols)
                    {
                        if (Colitem == "")
                        {
                            continue;
                        }
                        if (Colitem.ToLower().Contains("as"))
                        {
                            string Values = Colitem.Substring(0, Colitem.ToLower().IndexOf("as") - 1);
                            string names = Colitem.Substring(Colitem.ToLower().IndexOf("as") + 2);
                            if (tableitem.Columns.Contains(names) == false)
                            {
                                DataColumn newc = new DataColumn();
                                newc.ColumnName = names.Trim();


                                if (Values.ToLower().StartsWith("select"))
                                {
                                    Values = Values.Substring(Values.ToLower().IndexOf("select") + 6);
                                    Values = Values.Trim();
                                }

                                if (Values == "''")
                                {
                                    Values = "";
                                }
                                else if (Values.StartsWith("'"))
                                {
                                    Values = Values.Substring(1, Values.Length - 2);

                                }
                                if (Values == "1")
                                {
                                    newc.DataType = typeof(int);
                                }
                                else if (Values.StartsWith("0.0") || Values.StartsWith("(0.0"))
                                {
                                    newc.DataType = typeof(int);
                                }
                                newc.DefaultValue = (Values == "''" ? "" : Values);
                                tableitem.Columns.Add(newc);

                            }
                        }
                        else
                        {
                            if (tableitem.Columns.Contains(Colitem.Trim()) == false)
                            {
                                tableitem.Columns.Add(Colitem.Trim());
                            }
                        }
                    }
                    break;
                }
            }



            return result;

            #endregion




        }

        public static DataTable SelectFromXLS_POI(string ServerFileName, string sheetName, bool isFirstRowColumn)
        {
            IWorkbook workbook = null;
            FileStream fs = null;
            ISheet sheet = null;

            DataTable data = new DataTable();
            int startRow = 0;

            fs = new FileStream(ServerFileName, FileMode.Open, FileAccess.Read);
            if (ServerFileName.IndexOf(".xlsx") > 0) // 2007版本  
                workbook = new XSSFWorkbook(fs);
            else if (ServerFileName.IndexOf(".xls") > 0) // 2003版本  
                workbook = new HSSFWorkbook(fs);

            if (sheetName != null)
            {
                sheet = workbook.GetSheet(sheetName);
                if (sheet == null) //如果没有找到指定的sheetName对应的sheet，则尝试获取第一个sheet  
                {
                    sheet = workbook.GetSheetAt(0);
                }
            }
            else
            {
                sheet = workbook.GetSheetAt(0);
            }
            if (sheet != null)
            {
                IRow firstRow = sheet.GetRow(0);
                int cellCount = firstRow.LastCellNum; //一行最后一个cell的编号 即总的列数  
                if (cellCount == -1 && sheet.LastRowNum > 1)
                {
                    cellCount = sheet.GetRow(1).LastCellNum;
                }



                for (int i = 0; i < cellCount; ++i)
                {
                    ICell cell = firstRow.GetCell(i);
                    if (isFirstRowColumn)
                    {
                        if (cell != null)
                        {
                            string cellValue = cell.StringCellValue;


                            if (cellValue != null)
                            {
                                DataColumn column = new DataColumn(cellValue);
                                data.Columns.Add(column);
                            }
                            startRow = 1;

                        }





                    }
                    else
                    {
                        DataColumn column = new DataColumn("F" + i.ToString());
                        data.Columns.Add(column);
                        startRow = 0;
                    }
                }



                //最后一列的标号  
                int rowCount = sheet.LastRowNum;
                for (int i = startRow; i <= rowCount; ++i)
                {
                    IRow row = sheet.GetRow(i);
                    if (row == null || row.FirstCellNum == -1) continue; //没有数据的行默认是null　　　　　　　  

                    DataRow dataRow = data.NewRow();
                    for (int j = row.FirstCellNum; j < cellCount; ++j)
                    {
                        if (row.GetCell(j) != null) //同理，没有数据的单元格都默认是null  
                            dataRow[j] = row.GetCell(j).ToString();
                    }
                    data.Rows.Add(dataRow);
                }
            }
            fs.Close();
            return data;


        }
        public static DataTable SelectFromXLS_POIIndex(string ServerFileName, Int32 sheetindex, bool isFirstRowColumn)
        {
            IWorkbook workbook = null;
            FileStream fs = null;
            ISheet sheet = null;

            DataTable data = new DataTable();
            int startRow = 0;
            try
            {
                fs = new FileStream(ServerFileName, FileMode.Open, FileAccess.Read);
                if (ServerFileName.IndexOf(".xlsx") > 0) // 2007版本  
                    workbook = new XSSFWorkbook(fs);
                else if (ServerFileName.IndexOf(".xls") > 0) // 2003版本  
                    workbook = new HSSFWorkbook(fs);


                sheet = workbook.GetSheetAt(sheetindex);

                if (sheet != null)
                {
                    IRow firstRow = sheet.GetRow(0);
                    int cellCount = firstRow.LastCellNum; //一行最后一个cell的编号 即总的列数  




                    for (int i = firstRow.FirstCellNum; i < cellCount; ++i)
                    {
                        ICell cell = firstRow.GetCell(i);
                        if (cell != null)
                        {
                            string cellValue = cell.StringCellValue;
                            if (isFirstRowColumn)
                            {

                                if (cellValue != null)
                                {
                                    DataColumn column = new DataColumn(cellValue);
                                    data.Columns.Add(column);
                                }
                                startRow = 1;

                            }
                            else
                            {
                                DataColumn column = new DataColumn("F" + i.ToString());
                                data.Columns.Add(column);
                                startRow = 0;
                            }




                        }
                    }



                    //最后一列的标号  
                    int rowCount = sheet.LastRowNum;
                    for (int i = startRow; i <= rowCount; ++i)
                    {
                        IRow row = sheet.GetRow(i);
                        if (row == null) continue; //没有数据的行默认是null　　　　　　　  

                        DataRow dataRow = data.NewRow();
                        for (int j = row.FirstCellNum; j < cellCount; ++j)
                        {
                            if (row.GetCell(j) != null) //同理，没有数据的单元格都默认是null  
                                dataRow[j] = row.GetCell(j).ToString();
                        }
                        data.Rows.Add(dataRow);
                    }
                }

                return data;
            }
            catch (Exception ex)
            {
                Console.WriteLine("Exception: " + ex.Message,false);
                return null;
            }

        }





        public static void ExportToExcel(DataTable datasource, string sheetName, string NewFileName, string[] HeaderName)
        {

            //不允许dataGridView显示添加行，负责导出时会报最后一行未实例化错误

            HSSFWorkbook workbook = new HSSFWorkbook();
            ISheet sheet = workbook.CreateSheet(sheetName);
            IRow rowHead = sheet.CreateRow(0);

            //填写表头
            for (int i = 0; i < HeaderName.LongLength; i++)
            {
                rowHead.CreateCell(i, CellType.String).SetCellValue(HeaderName[i]);
            }
            //填写内容
            for (int i = 0; i < datasource.Rows.Count; i++)
            {
                IRow row = sheet.CreateRow(i + 1);
                for (int j = 0; j < datasource.Columns.Count; j++)
                {
                    row.CreateCell(j, CellType.String).SetCellValue(
                        datasource.Rows[i].Field<object>(j) == null ? "" : datasource.Rows[i].Field<object>(j).ToString()
                        );
                }
            }

            using (FileStream stream = File.OpenWrite(NewFileName))
            {
                workbook.Write(stream);
                stream.Close();
            }
            GC.Collect();
        }



        public static DataTable OpenCSV(string filePath)
        {
            Encoding encoding = System.Text.Encoding.Default; //Encoding.ASCII;//
            DataTable dt = new DataTable();
            FileStream fs = new FileStream(filePath, System.IO.FileMode.Open, System.IO.FileAccess.Read);

            //StreamReader sr = new StreamReader(fs, Encoding.UTF8);
            StreamReader sr = new StreamReader(fs, encoding);
            //string fileContent = sr.ReadToEnd();
            //encoding = sr.CurrentEncoding;
            //记录每次读取的一行记录
            string strLine = "";
            //记录每行记录中的各字段内容
            string[] aryLine = null;
            string[] tableHead = null;
            //标示列数
            int columnCount = 0;
            //标示是否是读取的第一行
            bool IsFirst = true;
            //逐行读取CSV中的数据
            while ((strLine = sr.ReadLine()) != null)
            {
                //strLine = Common.ConvertStringUTF8(strLine, encoding);
                //strLine = Common.ConvertStringUTF8(strLine);

                if (IsFirst == true)
                {
                    tableHead = strLine.Split(',');
                    IsFirst = false;
                    columnCount = tableHead.Length;
                    //创建列
                    for (int i = 0; i < columnCount; i++)
                    {
                        DataColumn dc = new DataColumn(tableHead[i]);
                        dt.Columns.Add(dc);
                    }
                }
                else
                {
                    aryLine = strLine.Split(',');
                    DataRow dr = dt.NewRow();
                    for (int j = 0; j < columnCount; j++)
                    {
                        dr[j] = aryLine[j];
                    }
                    dt.Rows.Add(dr);
                }
            }
            if (aryLine != null && aryLine.Length > 0)
            {
                dt.DefaultView.Sort = tableHead[0] + " " + "asc";
            }

            sr.Close();
            fs.Close();
            return dt;
        }


        public static DataSet SelectFromXLSNOHead(string ServerFileName, string SelectSQL)
        {
            string mystring = "Provider = Microsoft.Jet.OLEDB.4.0 ; Data Source = '" + ServerFileName + "';Extended Properties='Excel 8.0;HDR=No;IMEX=1'";
            OleDbConnection cnnxls = new OleDbConnection(mystring);
            OleDbDataAdapter myDa = null;
            DataSet myDs = new DataSet();
            try
            {
                cnnxls.Open();
                myDa = new OleDbDataAdapter(SelectSQL, cnnxls);
                myDa.Fill(myDs, "SelectResult");
            }
            catch (Exception AnyError)
            {
                cnnxls.Close();
                try
                {
                    return SelectFromXLS2007NoHead(ServerFileName, SelectSQL);
                }
                catch (Exception Error2)
                {

                    throw new Exception(SelectSQL + Error2.Message);
                }
                throw AnyError;
            }
            finally
            {
                cnnxls.Close();
            }
            return myDs;

        }

        public static DataSet SelectFromXLS2007NoHead(string ServerFileName, string SelectSQL)
        {
            string mystring = "Provider = Microsoft.ACE.OLEDB.12.0; Data Source = '" + ServerFileName + "';Extended Properties='Excel 12.0 Xml;HDR=No;IMEX=1'";
            OleDbConnection cnnxls = new OleDbConnection(mystring);
            OleDbDataAdapter myDa = null;
            DataSet myDs = new DataSet();
            try
            {
                cnnxls.Open();
                myDa = new OleDbDataAdapter(SelectSQL, cnnxls);
                myDa.Fill(myDs, "SelectResult");


            }
            catch (Exception AnyError)
            {
                cnnxls.Close();
                throw AnyError;
            }
            finally
            {
                cnnxls.Close();
            }
            return myDs;

        }


        /// <summary>
        /// 獲取工作表對應的SQL表名
        /// </summary>
        /// <param name="SheetName"></param>
        /// <returns></returns>
        public static string ConvertToSQLSheetName(string SheetName)
        {
            return "[" + SheetName + "$]";
        }
        /// <summary>
        /// 執行無返回查詢
        /// </summary>
        /// <param name="ServerFileName"></param>
        /// <param name="UpdateSql"></param>
        public static void ExcuteNonQuery(string ServerFileName, string UpdateSql)
        {
            string mystring = "Provider = Microsoft.Jet.OLEDB.4.0 ; Data Source = '" + ServerFileName + "';Extended Properties=Excel 8.0";
            OleDbConnection cnnxls = new OleDbConnection(mystring);
            OleDbCommand ToRun = new OleDbCommand(UpdateSql, cnnxls);
            try
            {
                cnnxls.Open();
                ToRun.ExecuteNonQuery();
            }
            catch (Exception AnyError)
            {
                cnnxls.Close();
                throw AnyError;
            }
            finally
            {
                cnnxls.Close();
            }
        }
        public static void ExcuteNonQuery2007(string ServerFileName, string UpdateSql)
        {
            string mystring = "Provider = Microsoft.ACE.OLEDB.12.0; Data Source = '" + ServerFileName + "';Extended Properties='Excel 12.0 Xml;HDR=No;'";
            OleDbConnection cnnxls = new OleDbConnection(mystring);
            OleDbCommand ToRun = new OleDbCommand(UpdateSql, cnnxls);
            try
            {
                cnnxls.Open();
                ToRun.ExecuteNonQuery();
            }
            catch (Exception AnyError)
            {
                cnnxls.Close();
                throw AnyError;
            }
            finally
            {
                cnnxls.Close();
            }
        }
        public static DataTable GetTables(string ServerFileName)
        {
            string mystring = "Provider = Microsoft.Jet.OLEDB.4.0 ; Data Source = '" + ServerFileName + "';Extended Properties=Excel 8.0";
            OleDbConnection cnnxls = new OleDbConnection(mystring);
            cnnxls.Open();
            DataTable dtOle = cnnxls.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "Table " });
            cnnxls.Close();
            return dtOle;
        }

        /// <summary>
        /// DataTable通过流导出Excel
        /// </summary>
        /// <param name="ds">数据源DataSet</param>
        /// <param name="columns">DataTable中列对应的列名(可以是中文),若为null则取DataTable中的字段名</param>
        /// <param name="fileName">保存文件名(例如：a.xls)</param>
        /// <returns></returns>
        public static bool StreamExport(DataTable dt, string[] columns, string fileName, System.Web.UI.Page pages)
        {
            if (dt.Rows.Count > 65535) //总行数大于Excel的行数 
            {
                throw new Exception("预导出的数据总行数大于excel的行数");
            }
            if (string.IsNullOrEmpty(fileName)) return false;

            StringBuilder content = new StringBuilder();
            StringBuilder strtitle = new StringBuilder();
            content.Append("<html xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:x='urn:schemas-microsoft-com:office:excel' xmlns='http://www.w3.org/TR/REC-html40'>");
            content.Append("<head><title></title><meta http-equiv='Content-Type' content=\"text/html; charset=gb2312\">");
            //注意：[if gte mso 9]到[endif]之间的代码，用于显示Excel的网格线，若不想显示Excel的网格线，可以去掉此代码
            content.Append("<!--[if gte mso 9]>");
            content.Append("<xml>");
            content.Append(" <x:ExcelWorkbook>");
            content.Append("  <x:ExcelWorksheets>");
            content.Append("   <x:ExcelWorksheet>");
            content.Append("    <x:Name>Sheet1</x:Name>");
            content.Append("    <x:WorksheetOptions>");
            content.Append("      <x:Print>");
            content.Append("       <x:ValidPrinterInfo />");
            content.Append("      </x:Print>");
            content.Append("    </x:WorksheetOptions>");
            content.Append("   </x:ExcelWorksheet>");
            content.Append("  </x:ExcelWorksheets>");
            content.Append("</x:ExcelWorkbook>");
            content.Append("</xml>");
            content.Append("<![endif]-->");
            content.Append("</head><body><table style='border-collapse:collapse;table-layout:fixed;'><tr>");

            if (columns != null)
            {
                for (int i = 0; i < columns.Length; i++)
                {
                    if (columns[i] != null && columns[i] != "")
                    {
                        content.Append("<td><b>" + columns[i] + "</b></td>");
                    }
                    else
                    {
                        content.Append("<td><b>" + dt.Columns[i].ColumnName + "</b></td>");
                    }
                }
            }
            else
            {
                for (int j = 0; j < dt.Columns.Count; j++)
                {
                    content.Append("<td><b>" + dt.Columns[j].ColumnName + "</b></td>");
                }
            }
            content.Append("</tr>\n");

            for (int j = 0; j < dt.Rows.Count; j++)
            {
                content.Append("<tr>");
                for (int k = 0; k < dt.Columns.Count; k++)
                {
                    object obj = dt.Rows[j][k];
                    Type type = obj.GetType();
                    if (type.Name == "Int32" || type.Name == "Single" || type.Name == "Double" || type.Name == "Decimal")
                    {
                        double d = obj == DBNull.Value ? 0.0d : Convert.ToDouble(obj);
                        if (type.Name == "Int32" || (d - Math.Truncate(d) == 0))
                            content.AppendFormat("<td style='vnd.ms-excel.numberformat:#,##0'>{0}</td>", obj);
                        else
                            content.AppendFormat("<td style='vnd.ms-excel.numberformat:#,##0.00'>{0}</td>", obj);
                    }
                    else
                        content.AppendFormat("<td style='vnd.ms-excel.numberformat:@'>{0}</td>", obj);
                }
                content.Append("</tr>\n");
            }
            content.Append("</table></body></html>");
            content.Replace("&nbsp;", "");
            pages.Response.Clear();
            pages.Response.Buffer = true;
            pages.Response.ContentType = "application/ms-excel";  //"application/ms-excel";
            pages.Response.Charset = "UTF-8";
            pages.Response.ContentEncoding = System.Text.Encoding.UTF7;
            fileName = System.Web.HttpUtility.UrlEncode(fileName, System.Text.Encoding.UTF8);
            pages.Response.AppendHeader("Content-Disposition", "attachment; filename=" + fileName);
            pages.Response.Write(content.ToString());
            //pages.Response.End();  //注意，若使用此代码结束响应可能会出现“由于代码已经过优化或者本机框架位于调用堆栈之上,无法计算表达式的值。”的异常。
            HttpContext.Current.ApplicationInstance.CompleteRequest(); //用此行代码代替上一行代码，则不会出现上面所说的异常。
            return true;
        }

        /// <summary>
        /// 将DataTable导出到Excel
        /// </summary>
        /// <param name="htmlTable">html表格内容</param> 
        /// <param name="fileName">仅文件名（非路径）</param> 
        /// <returns>返回Excel文件绝对路径</returns>
        public static string ExportHtmlTableToExcel(string htmlTable, string fileName)
        {
            string result;

            #region 第一步：将HtmlTable转换为DataTable
            htmlTable = htmlTable.Replace("\"", "'");
            var trReg = new Regex(pattern: @"(?<=(<[t|T][r|R]))[\s\S]*?(?=(</[t|T][r|R]>))");
            var trMatchCollection = trReg.Matches(htmlTable);
            DataTable dt = new DataTable("data");
            for (int i = 0; i < trMatchCollection.Count; i++)
            {
                var row = "<tr " + trMatchCollection[i].ToString().Trim() + "</tr>";
                var tdReg = new Regex(pattern: @"(?<=(<[t|T][d|D|h|H]))[\s\S]*?(?=(</[t|T][d|D|h|H]>))");
                var tdMatchCollection = tdReg.Matches(row);
                if (i == 0)
                {
                    foreach (var rd in tdMatchCollection)
                    {
                        var tdValue = RemoveHtml("<td " + rd.ToString().Trim() + "</td>");
                        DataColumn dc = new DataColumn(tdValue);
                        dt.Columns.Add(dc);
                    }
                }
                if (i > 0)
                {
                    DataRow dr = dt.NewRow();
                    for (int j = 0; j < tdMatchCollection.Count; j++)
                    {
                        var tdValue = RemoveHtml("<td " + tdMatchCollection[j].ToString().Trim() + "</td>");
                        dr[j] = tdValue;
                    }
                    dt.Rows.Add(dr);
                }
            }
            #endregion


            #region 第二步：将DataTable导出到Excel
            result = ExportDataSetToExcel(dt, fileName);
            #endregion


            return result;
        }


        /// <summary>
        /// 将DataTable导出到Excel
        /// </summary>
        /// <param name="dt">DataTable</param> 
        /// <param name="fileName">仅文件名（非路径）</param> 
        /// <returns>返回Excel文件绝对路径</returns>
        public static string ExportDataSetToExcel(DataTable dt, string fileName)
        {
            #region 表头
            XSSFWorkbook hssfworkbook = new XSSFWorkbook();
            NPOI.SS.UserModel.ISheet hssfSheet = hssfworkbook.CreateSheet("Data");
            hssfSheet.DefaultColumnWidth = 13;
            hssfSheet.SetColumnWidth(0, 25 * 256);
            hssfSheet.SetColumnWidth(3, 20 * 256);
            // 表头
            NPOI.SS.UserModel.IRow tagRow = hssfSheet.CreateRow(0);
            tagRow.Height = 22 * 20;


            // 标题样式
            NPOI.SS.UserModel.ICellStyle cellStyle = hssfworkbook.CreateCellStyle();
            cellStyle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
            cellStyle.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;
            cellStyle.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
            cellStyle.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
            cellStyle.BottomBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            cellStyle.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
            cellStyle.LeftBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            cellStyle.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
            cellStyle.RightBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            cellStyle.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
            cellStyle.TopBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;

            NPOI.SS.UserModel.ICellStyle numcellStyle = hssfworkbook.CreateCellStyle();
            numcellStyle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
            numcellStyle.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;
            numcellStyle.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
            numcellStyle.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
            numcellStyle.BottomBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            numcellStyle.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
            numcellStyle.LeftBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            numcellStyle.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
            numcellStyle.RightBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            numcellStyle.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
            numcellStyle.TopBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            NPOI.SS.UserModel.IDataFormat fmt = hssfworkbook.CreateDataFormat();
            numcellStyle.DataFormat = fmt.GetFormat("0.00");


            int colIndex;
            for (colIndex = 0; colIndex < dt.Columns.Count; colIndex++)
            {
                tagRow.CreateCell(colIndex).SetCellValue(dt.Columns[colIndex].ColumnName);
                tagRow.GetCell(colIndex).CellStyle = cellStyle;
            }
            #endregion
            #region 表数据
            // 表数据  
            for (int k = 0; k < dt.Rows.Count; k++)
            {
                DataRow dr = dt.Rows[k];
                NPOI.SS.UserModel.IRow row = hssfSheet.CreateRow(k + 1);
                for (int i = 0; i < dt.Columns.Count; i++)
                {




                    decimal tryd = 0;
                    bool IsDecimal = Decimal.TryParse(dr[i].ToString(), out tryd);

                    row.CreateCell(i);
                    if (IsDecimal)
                    {

                        row.GetCell(i).SetCellValue(Convert.ToDouble(dr[i]));
                    }

                    else
                    {
                        row.GetCell(i).SetCellValue(dr[i].ToString());
                    }


                }
            }
            #endregion
            FileStream file = new FileStream(HttpContext.Current.Request.PhysicalApplicationPath + "Upload/" + fileName + ".xlsx", FileMode.Create);
            hssfworkbook.Write(file);

            file.Close();
            var basePath = VirtualPathUtility.AppendTrailingSlash(HttpContext.Current.Request.ApplicationPath);
            return (basePath + "Upload/" + fileName + ".xlsx");
        }




        /// <summary>
        ///     去除HTML标记
        /// </summary>
        /// <param name="htmlstring"></param>
        /// <returns>已经去除后的文字</returns>
        public static string RemoveHtml(string htmlstring)
        {
            //删除脚本    
            htmlstring =
                Regex.Replace(htmlstring, @"<script[^>]*?>.*?</script>",
                              "", RegexOptions.IgnoreCase);
            //删除HTML    
            htmlstring = Regex.Replace(htmlstring, @"<(.[^>]*)>", "", RegexOptions.IgnoreCase);
            htmlstring = Regex.Replace(htmlstring, @"([\r\n])[\s]+", "", RegexOptions.IgnoreCase);
            htmlstring = Regex.Replace(htmlstring, @"-->", "", RegexOptions.IgnoreCase);
            htmlstring = Regex.Replace(htmlstring, @"<!--.*", "", RegexOptions.IgnoreCase);
            htmlstring = Regex.Replace(htmlstring, @"&(quot|#34);", "\"", RegexOptions.IgnoreCase);
            htmlstring = Regex.Replace(htmlstring, @"&(amp|#38);", "&", RegexOptions.IgnoreCase);
            htmlstring = Regex.Replace(htmlstring, @"&(lt|#60);", "<", RegexOptions.IgnoreCase);
            htmlstring = Regex.Replace(htmlstring, @"&(gt|#62);", ">", RegexOptions.IgnoreCase);
            htmlstring = Regex.Replace(htmlstring, @"&(nbsp|#160);", "   ", RegexOptions.IgnoreCase);
            htmlstring = Regex.Replace(htmlstring, @"&(iexcl|#161);", "\xa1", RegexOptions.IgnoreCase);
            htmlstring = Regex.Replace(htmlstring, @"&(cent|#162);", "\xa2", RegexOptions.IgnoreCase);
            htmlstring = Regex.Replace(htmlstring, @"&(pound|#163);", "\xa3", RegexOptions.IgnoreCase);
            htmlstring = Regex.Replace(htmlstring, @"&(copy|#169);", "\xa9", RegexOptions.IgnoreCase);
            htmlstring = Regex.Replace(htmlstring, @"&#(\d+);", "", RegexOptions.IgnoreCase);


            htmlstring = htmlstring.Replace("<", "");
            htmlstring = htmlstring.Replace(">", "");
            htmlstring = htmlstring.Replace("\r\n", "");
            return htmlstring;
        }

    }
    #endregion
    public class Util_File
    {
        public static string ReadToEnd(String FilePath, Encoding Enc)
        {
            FileStream fs = new FileStream(FilePath, FileMode.Open);
            byte[] buf = new byte[fs.Length];
            fs.Read(buf, 0, buf.Length);
            fs.Close();
            return Enc.GetString(buf);
        }
        public static void SaveToFile(String Content, String FilePath, Encoding Enc)
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
            fs.Write(bfs, 0, bfs.Length);
            fs.Flush();
            fs.Close();
        }
    }
    #region "請求工具"
    public partial class Util_Http
    {
        public static string GetAbosulateUrl()
        {
            HttpRequest AnyRequest = HttpContext.Current.Request;
            string Paramresult = AnyRequest.Url.AbsoluteUri.Replace(AnyRequest.AppRelativeCurrentExecutionFilePath.Replace("~/", ""), "");
            int questionPosition = Paramresult.LastIndexOf("?");
            if (questionPosition == -1)
            {
                return Paramresult;
            }
            else
            {
                return Paramresult.Substring(0, questionPosition);
            }
        }
        public static string GetDoMain()
        {
            string strurl = GetAbosulateUrl();
            string retval;
            string strregex = @"( .com/|.net/|.cn/|.org/|.gov/ )";
            Regex r = new Regex(strregex, RegexOptions.IgnoreCase);
            Match m = r.Match(strurl);
            retval = m.ToString();
            strregex = @".|/$";
            retval = Regex.Replace(retval, strregex, "").ToString();
            if (retval == "")
                retval = "other";
            return retval;

        }
        public static void PageAlert(Page AnyPage, string AlertMessage)
        {
            AnyPage.ClientScript.RegisterStartupScript(typeof(string), Guid.NewGuid().ToString(), "<script>alert(\"" + AlertMessage.Replace("\"", "").Replace("'", "").Replace(System.Environment.NewLine, "") + "\")</script>");
        }
        public static void PageAlertAndScript(Page AnyPage, string AlertMessage, string script)
        {
            AnyPage.ClientScript.RegisterStartupScript(typeof(string), Guid.NewGuid().ToString(), "<script>alert(\"" + AlertMessage.Replace("\"", "").Replace("'", "").Replace(System.Environment.NewLine, "") + "\");" + script + "</script>");
        }
        public static void PageAlertAndScript(Page AnyPage, string AlertMessage, bool Util_dhtmlclose)
        {
            if (Util_dhtmlclose == true)
            {
                AnyPage.ClientScript.RegisterStartupScript(typeof(string), Guid.NewGuid().ToString(), "<script>alert(\"" + AlertMessage.Replace("\"", "").Replace("'", "").Replace(System.Environment.NewLine, "") + "\");" + "parent.ModelJSFrameWindow.__doPostBack('Autorefresh','');;parent.showingwindow.close()" + "</script>");
            }
            else
            {
                AnyPage.ClientScript.RegisterStartupScript(typeof(string), Guid.NewGuid().ToString(), "<script>alert(\"" + AlertMessage.Replace("\"", "").Replace("'", "").Replace(System.Environment.NewLine, "") + "\");" + "</script>");

            }

        }
        public static void PageAlertAndGoTo(Page AnyPage, string AlertMessage, string GoToUrl)
        {
            AnyPage.ClientScript.RegisterStartupScript(typeof(string), "error", "<script>alert(\"" + AlertMessage + "\");location.href='" + GoToUrl + "'</script>");
        }
        public static void PageAlert(Page AnyPage, string AlertMessage, Boolean IsClose)
        {
            if (IsClose == false)
            {
                AnyPage.ClientScript.RegisterStartupScript(typeof(string), "error", "<script>alert(\"" + AlertMessage + "\")</script>");

            }
            else
            {
                AnyPage.ClientScript.RegisterStartupScript(typeof(string), "error", "<script>alert(\"" + AlertMessage + "\");window.close();</script>");

            }
        }
        public static void RunFunction(Page AnyPage, string Script)
        {
            AnyPage.ClientScript.RegisterStartupScript(typeof(string), "RunFuncton", "<script>" + Script + "</script>");
        }
        public static void RunFunction(Page AnyPage, string Script, string NickName)
        {
            AnyPage.ClientScript.RegisterStartupScript(typeof(string), NickName, "<script>" + Script + "</script>");
        }
        public static void PageShowModulDialog(Page AnyPage, string URL)
        {
            AnyPage.ClientScript.RegisterStartupScript(typeof(string), Guid.NewGuid().ToString(), "<script>showModalDialog(\"" + URL + "\")</script>");
        }

        public static string SafeRequest(string RequestName, HttpRequest AnyRequest, HttpServerUtility AnyServer)
        {
            if (AnyRequest.QueryString[RequestName] == null)
            {
                return "";
            }
            else
            {
                return AnyServer.UrlDecode(AnyRequest.QueryString[RequestName]);
            }
        }
        /// <summary>
        /// 上傳文件，重命名為GUID
        /// </summary>
        /// <param name="AnyUpload"></param>
        /// <param name="NewGuid"></param>
        /// <param name="ServerPath"></param>
        public static void SafeUpload(FileUpload AnyUpload, Guid NewGuid, string ServerPath)
        {
            try
            {
                AnyUpload.SaveAs(ServerPath + "\\" + NewGuid.ToString());
            }
            catch (Exception AnyError)
            {

                throw AnyError;
            }
        }
        /// <summary>
        /// 安全上傳，后綴自動加時間
        /// </summary>
        /// <param name="AnyUpload"></param>
        /// <param name="ServerPath"></param>
        public static string SafeUpload(FileUpload AnyUpload, string ServerPath)
        {
            try
            {
                string NewName = DateTime.Now.ToString("yyyyMMddhhmmssffff") + Guid.NewGuid().ToString() + AnyUpload.FileName;

                AnyUpload.SaveAs(ServerPath + "\\" + NewName);
                return NewName;

            }
            catch (Exception AnyError)
            {

                throw AnyError;
            }
        }
        public static string ReplaceUpload(FileUpload AnyUpload, string ServerPath, string ReplaceFileName)
        {
            try
            {
                File.Delete(ServerPath + "\\" + ReplaceFileName);
                AnyUpload.SaveAs(ServerPath + "\\" + ReplaceFileName);
                return ReplaceFileName;

            }
            catch (Exception AnyError)
            {

                throw AnyError;
            }
        }

        public static string SafeUploadReturnName(FileUpload AnyUpload, string ServerPath)
        {
            try
            {
                string NewName = DateTime.Now.ToShortDateString() + Guid.NewGuid().ToString() + AnyUpload.FileName;

                AnyUpload.SaveAs(ServerPath + "\\" + NewName);
                return NewName;

            }
            catch (Exception AnyError)
            {

                throw AnyError;
            }
        }
        /// <summary>
        /// 下載文件
        /// </summary>
        /// <param name="TargetFile"></param>
        /// <param name="AnyResponse"></param>
        /// <param name="RenameFileName"></param>
        public static void SafeDownLoad(FileInfo TargetFile, HttpResponse AnyResponse, string RenameFileName)
        {
            //AnyResponse.Clear();
            //AnyResponse.AppendHeader("Content-Disposition", "attachment;filename=" + HttpUtility.UrlEncode(RenameFileName, System.Text.Encoding.UTF8));
            //AnyResponse.AppendHeader("Content-Length", TargetFile.Length.ToString());
            //AnyResponse.ContentType = "application/octet-stream";
            //AnyResponse.WriteFile(TargetFile.FullName);


            long fileSize = TargetFile.Length;
            AnyResponse.Clear();

            AnyResponse.AddHeader("Content-Disposition", "attachement;filename=" + HttpUtility.UrlEncode(TargetFile.Name, System.Text.Encoding.UTF8));
            //指定文件大小   
            AnyResponse.AddHeader("Content-Length", fileSize.ToString());
            AnyResponse.ContentType = "application/octet-stream";
            AnyResponse.WriteFile(TargetFile.FullName, 0, fileSize);
            AnyResponse.Flush();

        }
        /// <summary>
        /// 輸出Loading畫面
        /// </summary>
        /// <param name="AnyResponse"></param>
        public static void PageStartWaiting()
        {
            HttpContext.Current.Response.Write(" <script language=JavaScript type=text/javascript>");
            HttpContext.Current.Response.Write("var t_id = setInterval(animate,20);");
            HttpContext.Current.Response.Write("var pos=0;var dir=2;var len=0;");
            HttpContext.Current.Response.Write("function animate(){");
            HttpContext.Current.Response.Write("var elem = document.getElementById('progress');");
            HttpContext.Current.Response.Write("if(elem != null) {");
            HttpContext.Current.Response.Write("if (pos==0) len += dir;");
            HttpContext.Current.Response.Write("if (len>32 || pos>79) pos += dir;");
            HttpContext.Current.Response.Write("if (pos>79) len -= dir;");
            HttpContext.Current.Response.Write(" if (pos>79 && len==0) pos=0;");
            HttpContext.Current.Response.Write("elem.style.left = pos;");
            HttpContext.Current.Response.Write("elem.style.width = len;");
            HttpContext.Current.Response.Write("}}");
            HttpContext.Current.Response.Write("function remove_loading() {");
            HttpContext.Current.Response.Write(" this.clearInterval(t_id);");
            HttpContext.Current.Response.Write("var targelem = document.getElementById('loader_container');");
            HttpContext.Current.Response.Write("targelem.style.display='none';");
            HttpContext.Current.Response.Write("targelem.style.visibility='hidden';");
            HttpContext.Current.Response.Write("}");
            HttpContext.Current.Response.Write("</script>");
            HttpContext.Current.Response.Write("<style>");
            HttpContext.Current.Response.Write("#loader_container {text-align:center; position:absolute; top:40%; width:100%; left: 0;}");
            HttpContext.Current.Response.Write("#loader {font-family:Tahoma, Helvetica, sans; font-size:11.5px; color:#000000; background-color:#FFFFFF; padding:10px 0 16px 0; margin:0 auto; display:block; width:130px; border:1px solid #5a667b; text-align:left; z-index:2;}");
            HttpContext.Current.Response.Write("#progress {height:5px; font-size:1px; width:1px; position:relative; top:1px; left:0px; background-color:#8894a8;}");
            HttpContext.Current.Response.Write("#loader_bg {background-color:#e4e7eb; position:relative; top:8px; left:8px; height:7px; width:113px; font-size:1px;}");
            HttpContext.Current.Response.Write("</style>");
            HttpContext.Current.Response.Write("<div id=loader_container>");
            HttpContext.Current.Response.Write("<div id=loader>");
            HttpContext.Current.Response.Write("<div id='ProgressContent' align=center>頁面正在加載中 </div>");
            HttpContext.Current.Response.Write("<div id=loader_bg><div id=progress> </div></div>");
            HttpContext.Current.Response.Write("</div></div>");
            HttpContext.Current.Response.Flush();
        }

        public static void PageChangeValue(string Infomation)
        {
            HttpContext.Current.Response.Write(" <script language=JavaScript type=text/javascript>");
            HttpContext.Current.Response.Write("document.getElementById('ProgressContent').innerHTML='" + Infomation.Replace("'", "").Replace("\r\n", "") + "';");

            HttpContext.Current.Response.Write(" </script>");
            HttpContext.Current.Response.Flush();

        }
        /// <summary>
        /// 結束Loading畫面
        /// </summary>
        /// <param name="AnyResponse"></param>
        /// 
        public static void PageEndWaiting()
        {
            HttpContext.Current.Response.Write(" <script language=JavaScript type=text/javascript>");
            HttpContext.Current.Response.Write("remove_loading();");
            HttpContext.Current.Response.Write(" </script>");
            HttpContext.Current.Response.Flush();
        }
        public static void PageDownLoad(string FileName)
        {
            HttpContext.Current.Response.AddHeader("Content-Disposition", "attachment;filename=" + FileName);
            HttpContext.Current.Response.ContentType = "application/octet-stream";

        }

        public static string UrlEncode(HttpServerUtility AnyServer, string ToEncode)
        {
            return AnyServer.UrlEncode(ToEncode);
        }
        public static string UrlEncode(string ToEncode)
        {
            return HttpUtility.UrlEncode(ToEncode);
        }

        public static void SaveRequestToFile(string FilePath)
        {
            HttpRequest CurrentRequest = HttpContext.Current.Request;
            if (File.Exists(FilePath))
            {
                File.Delete(FilePath);
            }
            FileStream WriteStream = new FileStream(FilePath, FileMode.CreateNew);
            Stream Read = CurrentRequest.InputStream;
            byte[] TempByte = new byte[255];
            int HaveRead = Read.Read(TempByte, 0, 255);
            while (HaveRead == 255)
            {
                WriteStream.Write(TempByte, 0, 255);
                HaveRead = Read.Read(TempByte, 0, 255);
            }
            WriteStream.Write(TempByte, 0, HaveRead);
            WriteStream.Close();

        }

        public static string GetHttpContent(string Url, Encoding Enc)
        {
            WebClient LoadPage = new WebClient();
            Byte[] pageData = LoadPage.DownloadData(Url);//从指定网站下载数据
            return Enc.GetString(pageData);
        }
        public static string GetHttpContent(string Url, NameValueCollection QueryString, Encoding Enc)
        {
            WebClient LoadPage = new WebClient();
            LoadPage.QueryString = QueryString;
            Byte[] pageData = LoadPage.DownloadData(Url);//从指定网站下载数据
            return Enc.GetString(pageData);
        }
        public static string Url_AutoConvert(string Url)
        {
            if (Url.StartsWith("http://"))
            {
                return Url;
            }
            else
            {
                return Util_Http.GetAbosulateUrl() + Url.Replace("../", "");
            }
        }
        public static void HttpDownload(string url, string filename)
        {
            if (File.Exists(filename))
            {
                File.Delete(filename);
            }
            try
            {
                WebClient download = new WebClient();
                download.DownloadFile(url, filename);
            }
            catch (Exception)
            {

            }

        }



    }
    public class PostParam
    {
        public string Name { get { return _Name; } set { _Name = value; } }
        public string Data { get { return _Data; } set { _Data = value; } }
        private string _Name;
        private string _Data;
    }
    #endregion
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

        public static string SerializeDataTableXml(DataTable pDt)
        {
            // 序列化DataTable
            StringBuilder sb = new StringBuilder();
            XmlWriter writer = XmlWriter.Create(sb);
            XmlSerializer serializer = new XmlSerializer(typeof(DataTable));
            serializer.Serialize(writer, pDt);
            writer.Close();

            return sb.ToString();
        }

        /// <summary>
        /// 反序列化DataTable
        /// </summary>
        /// <param name="pXml">序列化的DataTable</param>
        /// <returns>DataTable</returns>
        public static DataTable DeserializeDataTableXml(string pXml)
        {

            StringReader strReader = new StringReader(pXml);
            XmlReader xmlReader = XmlReader.Create(strReader);
            XmlSerializer serializer = new XmlSerializer(typeof(DataTable));

            DataTable dt = serializer.Deserialize(xmlReader) as DataTable;

            return dt;
        }

        //Base64 序列化

        //        DataTable dt = new DataTable(); //用来转成byte[]的实例
        //        dt.Columns.Add("a");
        //dt.Rows.Add("b"); //添加一条测试数据 b
        //System.IO.MemoryStream memory = new MemoryStream();//使用内存流来存这些byte[]
        //        BinaryFormatter b = new BinaryFormatter();
        //        b.Serialize(memory,dt); //系列化datatable,MS已经对datatable实现了系列化接口,如果你自定义的类要系列化,实现IFormatter 就可以类似做法
        //byte[] buff = memory.GetBuffer(); //这里就可你想要的byte[],可以使用它来传输
        //        memory.Close();

        ////假如接收的仍是这个byte[] buff,这样来反系列化

        //DataTable dt1 = (DataTable)b.Deserialize(new MemoryStream(buff)); //dt1是byte[]转回的datatable
        //        Response.Write(dt1.Rows[0][0].ToString());

        ////输出的是 "b"

    }

    /// <summary>
    /// 这个类要尽可能保存，频繁创建会报调用频繁错误
    /// </summary>
    public class Util_WeChatEnterpriseMsg
    {
        public Util_WeChatEnterpriseMsg(string _Enterpriseid, string _AppSecret, string _AppAgent)
        {
            Enterpriseid = _Enterpriseid;
            AppSecret = _AppSecret;
            AppAgent = _AppAgent;
        }
        public string Enterpriseid
        {
            get;
            set;
        }
        public string AppSecret { get; set; }
        public string AppAgent { get; set; }
        private CookieCollection cookie = new CookieCollection();
        public String AccessToken
        {
            //企业ID            wx48e213f50ad641c8

            //请求方式： GET（HTTPS）
            //请求地址： https://qyapi.weixin.qq.com/cgi-bin/gettoken?corpid=ID&corpsecret=SECRET


            //AgentId            1000005
            //Secret RysOH8IiVWHvi5kWyyheee7YAlvE6Z4q9uDRvHrYxqI
            get
            {
                if (AccessTokenTimeOut == null || (DateTime.Now - AccessTokenTimeOut.Value).TotalMinutes > -5)
                {
                    String httpr = Util_WEB.OpenUrl("https://qyapi.weixin.qq.com/cgi-bin/gettoken?corpid=" + Enterpriseid + "&corpsecret=" + AppSecret, "", "", "GET", cookie);
                    NetFramework.Console.Write(httpr);
                    Newtonsoft.Json.Linq.JObject jr = Newtonsoft.Json.Linq.JObject.Parse(httpr);
                    if (jr["errcode"].ToString() == "0")
                    {
                        _AccessToken = jr["access_token"].ToString();
                        AccessTokenTimeOut = DateTime.Now.AddSeconds(Convert.ToInt32(jr["expires_in"].ToString()));
                        return _AccessToken;
                    }
                    else
                    {
                        return "get accesstoken error";
                    }
                }
                else
                {
                    return _AccessToken;
                }
            }

        }
        private String _AccessToken;
        public DateTime? AccessTokenTimeOut { get; set; }

        public string SendTextMsg(string HtmlContent, string touser = "@all")
        {
            // https://qyapi.weixin.qq.com/cgi-bin/message/send?access_token=ACCESS_TOKEN
            //           {
            //               "touser" : "UserID1|UserID2|UserID3",
            //  "toparty" : "PartyID1|PartyID2",
            //  "totag" : "TagID1 | TagID2",
            //  "msgtype" : "text",
            //  "agentid" : 1,
            //  "text" : {
            //                   "content" : "你的快递已到，请携带工卡前往邮件中心领取。\n出发前可查看<a href=\"http://work.weixin.qq.com\">邮件中心视频实况</a>，聪明避开排队。"
            //  },
            //  "safe":0,
            //  "enable_id_trans": 0,
            //  "enable_duplicate_check": 0,
            //  "duplicate_check_interval": 1800
            //}
            //            参数说明：

            //参数 是否必须    说明
            //touser  否 指定接收消息的成员，成员ID列表（多个接收者用‘|’分隔，最多支持1000个）。
            //特殊情况：指定为”@all”，则向该企业应用的全部成员发送
            //toparty 否 指定接收消息的部门，部门ID列表，多个接收者用‘|’分隔，最多支持100个。
            //当touser为”@all”时忽略本参数
            //totag   否 指定接收消息的标签，标签ID列表，多个接收者用‘|’分隔，最多支持100个。
            //当touser为”@all”时忽略本参数
            //msgtype 是 消息类型，此时固定为：text
            //agentid 是 企业应用的id，整型。企业内部开发，可在应用的设置页面查看；第三方服务商，可通过接口 获取企业授权信息 获取该参数值
            //content 是 消息内容，最长不超过2048个字节，超过将截断（支持id转译）
            //safe 否   表示是否是保密消息，0表示否，1表示是，默认0
            //enable_id_trans 否 表示是否开启id转译，0表示否，1表示是，默认0
            //enable_duplicate_check  否 表示是否开启重复消息检查，0表示否，1表示是，默认0
            //duplicate_check_interval    否 表示是否重复消息检查的时间间隔，默认1800s，最大不超过4小时
            Newtonsoft.Json.Linq.JObject BODY = new Newtonsoft.Json.Linq.JObject();

            BODY["touser"] = touser;
            BODY["toparty"] = "";
            BODY["msgtype"] = "text";
            BODY["agentid"] = AppAgent;


            Newtonsoft.Json.Linq.JObject text = new Newtonsoft.Json.Linq.JObject();
            text["content"] = HtmlContent;
            BODY["text"] = text;

            BODY["safe"] = 0;
            BODY["enable_id_trans"] = 0;
            BODY["enable_duplicate_check"] = 0;
            BODY["duplicate_check_interval"] = 1800;
            CookieCollection cookie = new CookieCollection();

            String Return = NetFramework.Util_WEB.OpenUrl("https://qyapi.weixin.qq.com/cgi-bin/message/send?access_token=" + AccessToken, "", BODY.ToString(), "POST", cookie);
            NetFramework.Console.Write(Return);
            return Return;
        }

    }

    public class Util_CMD
    {
        // <summary>
        /// 运行cmd命令
        /// 会显示命令窗口
        /// </summary>
        /// <param name="cmdExe">指定应用程序的完整路径</param>
        /// <param name="cmdStr">执行命令行参数</param>
        public static bool RunCmd(string cmdExe, string cmdStr)
        {
            bool result = false;
            try
            {
                using (Process myPro = new Process())
                {
                    //指定启动进程是调用的应用程序和命令行参数
                    ProcessStartInfo psi = new ProcessStartInfo(cmdExe, cmdStr);
                    myPro.StartInfo = psi;
                    myPro.Start();
                    myPro.WaitForExit();
                    result = true;
                }
            }
            catch
            {

            }
            return result;
        }

        /// <summary>
        /// 运行cmd命令
        /// 不显示命令窗口
        /// </summary>
        /// <param name="cmdExe">指定应用程序的完整路径</param>
        /// <param name="cmdStr">执行命令行参数</param>
        public static bool RunHideCmd(string cmdExe, string cmdStr)
        {
            bool result = false;
            try
            {
                using (Process myPro = new Process())
                {
                    myPro.StartInfo.FileName = "cmd.exe";
                    myPro.StartInfo.UseShellExecute = false;
                    myPro.StartInfo.RedirectStandardInput = true;
                    myPro.StartInfo.RedirectStandardOutput = true;
                    myPro.StartInfo.RedirectStandardError = true;
                    myPro.StartInfo.CreateNoWindow = true;
                    myPro.Start();
                    //如果调用程序路径中有空格时，cmd命令执行失败，可以用双引号括起来 ，在这里两个引号表示一个引号（转义）
                    string str = string.Format(@"""{0}"" {1} {2}", cmdExe, cmdStr, "&exit");

                    myPro.StandardInput.WriteLine(str);
                    myPro.StandardInput.AutoFlush = true;
                    myPro.WaitForExit();

                    result = true;
                }
            }
            catch
            {

            }
            return result;
        }

    }

    #region"XML工具"
    public partial class Util_XML
    {
        public static string GetAttribute(string XMLData, string ToGetAttribute)
        {
            XmlDocument ToLoad = new XmlDocument();
            ToLoad.LoadXml(XMLData);
            return ToLoad.DocumentElement.GetAttribute(ToGetAttribute);
        }

        public static DataTable GetPathDataTable(string XMLData, string Path)
        {
            XmlDocument ToLoad = new XmlDocument();
            ToLoad.LoadXml(XMLData);
            XmlNodeList ResultList = ToLoad.SelectNodes(Path);
            DataTable Result = new DataTable();
            Result.Columns.Add(new DataColumn("XmlData", typeof(string)));
            foreach (XmlNode item in ResultList)
            {
                Result.Rows.Add(item.OuterXml);
            }
            return Result;
        }
        public static DataTable GetPathDataTable(XmlDocument XMLDoc, string Path)
        {
            XmlNodeList ResultList = XMLDoc.SelectNodes(Path);
            DataTable Result = new DataTable();
            Result.Columns.Add(new DataColumn("XmlData", typeof(string)));
            foreach (XmlNode item in ResultList)
            {
                Result.Rows.Add(item.OuterXml);
            }
            return Result;
        }
        public static XmlDocument LoadDoc(string XMLData)
        {
            XmlDocument ToReturn = new XmlDocument();
            ToReturn.LoadXml(XMLData);
            return ToReturn;
        }
        public XmlNodeList GetPathNodeList(XmlDocument XMLDoc, string Path)
        {
            return XMLDoc.SelectNodes(Path);

        }
        public XmlNodeList GetPathNodeList(string XMLData, string Path)
        {
            XmlDocument ToLoad = new XmlDocument();
            ToLoad.LoadXml(XMLData);
            return ToLoad.SelectNodes(Path);
        }
        public static void JoinXml(ref XmlDocument XmlDoc, string JoinData)
        {
            XmlNode TheFragument = XmlDoc.CreateNode(XmlNodeType.DocumentFragment, "Nothing", "");
            TheFragument.InnerXml = JoinData;
            XmlDoc.DocumentElement.AppendChild(TheFragument);

        }
        public static void SetArtibuteWithNull(ref XmlNode DataNode, string ArrtibuteName, string Value, bool EmptyToNull)
        {
            if (EmptyToNull == true)
            {
                if (Value == "")
                {
                    ((XmlElement)DataNode).RemoveAttribute(ArrtibuteName);
                }
                else
                {
                    ((XmlElement)DataNode).SetAttribute(ArrtibuteName, Value);
                }
            }
            else
            {
                ((XmlElement)DataNode).SetAttribute(ArrtibuteName, Value);
            }
        }
        public static void SetArtibuteWithNull(XmlElement DataNode, string ArrtibuteName, string Value, bool EmptyToNull)
        {
            if (EmptyToNull == true)
            {
                if (Value == "")
                {
                    ((XmlElement)DataNode).RemoveAttribute(ArrtibuteName);
                }
                else
                {
                    ((XmlElement)DataNode).SetAttribute(ArrtibuteName, Value);
                }
            }
            else
            {
                ((XmlElement)DataNode).SetAttribute(ArrtibuteName, Value);
            }
        }

    }

    #endregion

    public class Util_Datetime
    {
        public static string GetDayOfWeek(DayOfWeek ToGet)
        {
            switch (ToGet)
            {
                case DayOfWeek.Friday:
                    return "星期五";
                case DayOfWeek.Monday:
                    return "星期一";
                case DayOfWeek.Saturday:
                    return "星期六";
                case DayOfWeek.Sunday:
                    return "星期天";
                case DayOfWeek.Thursday:
                    return "星期四";
                case DayOfWeek.Tuesday:
                    return "星期二";
                case DayOfWeek.Wednesday:
                    return "星期三";
                default:
                    return "未定义";
            }
        }
        public static string GetSlimDayOfWeek(DayOfWeek ToGet)
        {
            switch (ToGet)
            {
                case DayOfWeek.Friday:
                    return "五";
                case DayOfWeek.Monday:
                    return "一";
                case DayOfWeek.Saturday:
                    return "六";
                case DayOfWeek.Sunday:
                    return "日";
                case DayOfWeek.Thursday:
                    return "四";
                case DayOfWeek.Tuesday:
                    return "二";
                case DayOfWeek.Wednesday:
                    return "三";
                default:
                    return "未定义";

            }
        }
    }

    public class Util_Thumbmail
    {
        public static void GenThumbnail(string pathImageFrom, string pathImageTo, int width, int height)
        {
            System.Drawing.Image imageFrom = null;
            try
            {
                imageFrom = System.Drawing.Image.FromFile(pathImageFrom);
            }
            catch (Exception AnyError)
            {
                throw AnyError;
            }
            if (imageFrom == null)
            {
                return;
            }
            // 源图宽度及高度 
            int imageFromWidth = imageFrom.Width;
            int imageFromHeight = imageFrom.Height;
            // 生成的缩略图实际宽度及高度 
            int bitmapWidth = width;
            int bitmapHeight = height;
            // 生成的缩略图在上述"画布"上的位置 
            int X = 0;
            int Y = 0;
            // 根据源图及欲生成的缩略图尺寸,计算缩略图的实际尺寸及其在"画布"上的位置 
            if (bitmapHeight * imageFromWidth > bitmapWidth * imageFromHeight)
            {
                bitmapHeight = imageFromHeight * width / imageFromWidth;
                Y = (height - bitmapHeight) / 2;
            }
            else
            {
                bitmapWidth = imageFromWidth * height / imageFromHeight;
                X = (width - bitmapWidth) / 2;
            }
            // 创建画布 
            Bitmap bmp = new Bitmap(width, height);
            Graphics g = Graphics.FromImage(bmp);
            // 用白色清空 
            g.Clear(Color.White);
            // 指定高质量的双三次插值法。执行预筛选以确保高质量的收缩。此模式可产生质量最高的转换图像。 
            g.InterpolationMode = InterpolationMode.HighQualityBicubic;
            // 指定高质量、低速度呈现。 
            g.SmoothingMode = SmoothingMode.HighQuality;
            // 在指定位置并且按指定大小绘制指定的 Image 的指定部分。 
            g.DrawImage(imageFrom, new Rectangle(X, Y, bitmapWidth, bitmapHeight), new Rectangle(0, 0, imageFromWidth, imageFromHeight), GraphicsUnit.Pixel);
            try
            {
                //经测试 .jpg 格式缩略图大小与质量等最优 
                bmp.Save(pathImageTo, ImageFormat.Jpeg);
            }
            catch
            {
            }
            finally
            {
                //显示释放资源 
                imageFrom.Dispose();
                bmp.Dispose();
                g.Dispose();
            }
        }
        public static void ScaleAndCut(string pathImageFrom, int ScaleWidth, int ScaleHeight, int CutWidth, int CutHeight, string pathImageTo)
        {
            System.Drawing.Image imageFrom = null;
            try
            {
                imageFrom = System.Drawing.Image.FromFile(pathImageFrom);
            }
            catch (Exception AnyError)
            {
                throw AnyError;
            }
            if (imageFrom == null)
            {
                return;
            }
            Bitmap bmp = new Bitmap(CutWidth, CutHeight);
            Graphics g = Graphics.FromImage(bmp);
            g.Clear(Color.White);
            // 指定高质量的双三次插值法。执行预筛选以确保高质量的收缩。此模式可产生质量最高的转换图像。 
            g.InterpolationMode = InterpolationMode.HighQualityBicubic;
            // 指定高质量、低速度呈现。 
            g.SmoothingMode = SmoothingMode.HighQuality;
            // 在指定位置并且按指定大小绘制指定的 Image 的指定部分。 
            g.DrawImage(imageFrom, new Rectangle(0, 0, CutWidth, CutHeight), new Rectangle((imageFrom.Width - imageFrom.Height) / 2, 0, imageFrom.Height, imageFrom.Height), GraphicsUnit.Pixel);

            try
            {
                bmp.Save(pathImageTo, ImageFormat.Jpeg);
            }
            catch (Exception)
            {

                throw;
            }
            finally
            {
                bmp.Dispose();
                g.Dispose();
            }
        }
        public static System.Drawing.Image ThumbView(string path, int width, int height)
        {
            try
            {


                System.Drawing.Image Result = new System.Drawing.Bitmap(width, height);

                System.Drawing.Image imageFrom = System.Drawing.Image.FromFile(path);
                Graphics g = Graphics.FromImage(Result);
                // 用白色清空 
                g.Clear(Color.White);
                // 指定高质量的双三次插值法。执行预筛选以确保高质量的收缩。此模式可产生质量最高的转换图像。 
                g.InterpolationMode = InterpolationMode.HighQualityBicubic;

                // 指定高质量、低速度呈现。 
                g.SmoothingMode = SmoothingMode.HighQuality;
                // 在指定位置并且按指定大小绘制指定的 Image 的指定部分。 
                g.DrawImage(imageFrom, new Rectangle(0, 0, width, height), new Rectangle(0, 0, imageFrom.Width, imageFrom.Height), GraphicsUnit.Pixel);

                imageFrom.Dispose();
                g.Dispose();
                return Result;
            }
            catch (Exception)
            {
                return null;

            }
        }
        public static System.Drawing.Image RightHalfCutThumbView(string path, int width, int height)
        {
            try
            {


                System.Drawing.Image Result = new System.Drawing.Bitmap(width, height);

                System.Drawing.Image imageFrom = System.Drawing.Image.FromFile(path);
                Graphics g = Graphics.FromImage(Result);
                // 用白色清空 
                g.Clear(Color.White);
                // 指定高质量的双三次插值法。执行预筛选以确保高质量的收缩。此模式可产生质量最高的转换图像。 
                g.InterpolationMode = InterpolationMode.HighQualityBicubic;

                // 指定高质量、低速度呈现。 
                g.SmoothingMode = SmoothingMode.HighQuality;
                // 在指定位置并且按指定大小绘制指定的 Image 的指定部分。 
                g.DrawImage(imageFrom, new Rectangle(0, 0, width / 2, height), new Rectangle(imageFrom.Width / 2, 0, imageFrom.Width / 2, imageFrom.Height / 2), GraphicsUnit.Pixel);
                g.DrawImage(imageFrom, new Rectangle(width / 2, 0, width, height), new Rectangle(imageFrom.Width / 2, imageFrom.Height / 2, imageFrom.Width / 2, imageFrom.Height / 2), GraphicsUnit.Pixel);

                imageFrom.Dispose();
                g.Dispose();
                return Result;
            }
            catch (Exception)
            {
                return null;

            }
        }

        public static Bitmap BitmapConvetGray(Bitmap img)
        {

            int h = img.Height;

            int w = img.Width;

            int gray = 0;    //灰度值

            Bitmap bmpOut = new Bitmap(w, h, PixelFormat.Format24bppRgb);    //每像素3字节

            BitmapData dataIn = img.LockBits(new Rectangle(0, 0, w, h), ImageLockMode.ReadOnly, PixelFormat.Format24bppRgb);

            BitmapData dataOut = bmpOut.LockBits(new Rectangle(0, 0, w, h), ImageLockMode.ReadWrite, PixelFormat.Format24bppRgb);

            unsafe
            {

                byte* pIn = (byte*)(dataIn.Scan0.ToPointer());      //指向源文件首地址

                byte* pOut = (byte*)(dataOut.Scan0.ToPointer());  //指向目标文件首地址

                for (int y = 0; y < dataIn.Height; y++)  //列扫描
                {

                    for (int x = 0; x < dataIn.Width; x++)   //行扫描
                    {

                        gray = (pIn[0] * 19595 + pIn[1] * 38469 + pIn[2] * 7472) >> 16;  //灰度计算公式

                        pOut[0] = (byte)gray;     //R分量

                        pOut[1] = (byte)gray;     //G分量

                        pOut[2] = (byte)gray;     //B分量

                        pIn += 3; pOut += 3;      //指针后移3个分量位置

                    }

                    pIn += dataIn.Stride - dataIn.Width * 3;

                    pOut += dataOut.Stride - dataOut.Width * 3;

                }

            }

            bmpOut.UnlockBits(dataOut);

            img.UnlockBits(dataIn);

            return bmpOut;

        }
        public static System.Drawing.Image ByteToImage(byte[] Result)
        {
            MemoryStream ms = new System.IO.MemoryStream(Result);
            System.Drawing.Image img = System.Drawing.Image.FromStream(ms);
            return img;
        }
    }

}

