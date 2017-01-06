using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net;
using System.IO;
using System.Security.Cryptography.X509Certificates;
using System.Net.Security;
using System.Windows;
using System.Text.RegularExpressions;
using System.Web;
namespace BIADBIMnavisworks
{
    class HttpWebResponseUtility
    {
        private static readonly string DefaultUserAgent = "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.2; SV1; .NET CLR 1.1.4322; .NET CLR 2.0.50727)";//浏览器  
        private static Encoding requestEncoding = System.Text.Encoding.UTF8;//字符集

        //public static string tagUrl = "http://172.18.83.118/obix/config/ObixPointTest/";
        public static string tagUrl = "";
        public static string tagUrl1 = "http://123.127.216.104:8088/OBIXtransmit/WebServiceOBIXTransmit.asmx";

        public static string userName = "admin";
        public static string password = "";

        public static string authData = "";

        public static int timeout = 1000;



        public static string CreateGetHttpResponse()
        {
            string responseFromServer = "";

            try
            {
                HttpWebRequest request = WebRequest.Create(tagUrl) as HttpWebRequest;
                request.Timeout = timeout;
                request.Credentials = new NetworkCredential(userName, password);

                authData = userName + ":" + password;

                byte[] encData_byte = new byte[authData.Length];
                encData_byte = System.Text.Encoding.UTF8.GetBytes(authData);
                //request.Headers.Add("Authorization: Basic " + Convert.ToBase64String(encData_byte));

                //读取返回
                HttpWebResponse response = (HttpWebResponse)request.GetResponse();
                //使用Stream操作返回数据
                Stream dataStream = response.GetResponseStream();
                //封装为StreamReader，操作更加简单
                StreamReader reader = new StreamReader(dataStream);
                //读取内容
                responseFromServer = reader.ReadToEnd();
                //清理工作
                reader.Close();
                dataStream.Close();
                response.Close();
            }
            catch
            {
                return "";
            }

            return responseFromServer;
        }

        //分析字符串取值
        public static string GetResultObix(string sObixR)
        {
            if (sObixR == "")
                return "";

            return sObixR.Split(' ')[0];
        }

        public static string SOAPAction = "";
        //读取转发
        public static string CreateGetHttpResponseTransmit() //由tagUrl1转发
        {

            //构造soap请求信息
            StringBuilder soap = new StringBuilder();
            string responseFromServer = "";


            try
            {
                soap.Append("<?xml version=\"1.0\" encoding=\"utf-8\"?>");
                soap.Append("<soap:Envelope xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xmlns:xsd=\"http://www.w3.org/2001/XMLSchema\" xmlns:soap=\"http://schemas.xmlsoap.org/soap/envelope/\">");
                soap.Append("<soap:Body>");

                soap.Append("<OBIXtransmit xmlns=\"http://tempuri.org/\">");
                soap.Append("<URL>" + tagUrl.Trim() + "</URL>");
                soap.Append("</OBIXtransmit>");

                //soap.Append("<HelloWorld xmlns=\"http://tempuri.org/\" />");
                soap.Append("</soap:Body>");
                soap.Append("</soap:Envelope>");

                //发起请求  
                Uri uri = new Uri(tagUrl1);
                WebRequest webRequest = WebRequest.Create(uri);
                webRequest.Headers.Add("SOAPAction", "http://tempuri.org/OBIXtransmit");
                webRequest.Timeout = 10000;
                webRequest.ContentType = "text/xml; charset=GB2312";
                webRequest.Method = "POST";
                using (Stream requestStream = webRequest.GetRequestStream())
                {
                    byte[] paramBytes = Encoding.UTF8.GetBytes(soap.ToString());
                    requestStream.Write(paramBytes, 0, paramBytes.Length);
                }

                WebResponse webResponse = webRequest.GetResponse();
                StreamReader reader = new StreamReader(webResponse.GetResponseStream());
                responseFromServer = reader.ReadToEnd();

                //整理
                responseFromServer = responseFromServer.Substring(responseFromServer.IndexOf("<OBIXtransmitResult>") + "<OBIXtransmitResult>".Length);
                responseFromServer = responseFromServer.Substring(0, responseFromServer.IndexOf("</OBIXtransmitResponse>") - "</OBIXtransmitResponse>".Length+1);
                //responseFromServer = responseFromServer.Replace("&lt;", "<");
                //responseFromServer = responseFromServer.Replace("&gt;", ">");

                responseFromServer = HttpUtility.HtmlDecode(responseFromServer);
            }
            catch
            {

            }
            
            return responseFromServer;
        }
    }
}
