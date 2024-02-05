using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Net.Security;
using System.Net;
using System.Security.Cryptography.X509Certificates;
using System.ServiceProcess;
using System.Text;
using System.Threading;
using System.Net.Mail;

namespace eNotieService
{
    public partial class Service1 : ServiceBase
    {
        private Timer Timer;
        private int TimerLag = Convert.ToInt16(ConfigurationManager.AppSettings["TimerLag"].Trim());
        private int TimerPeriod = Convert.ToInt16(ConfigurationManager.AppSettings["TimerPeriod"].Trim());
        public Service1()
        {
            InitializeComponent();

        }

        protected override void OnStart(string[] args)
        {
            try
            {
                PrintLog("Service Started at " + DateTime.Now.ToString(), "1", "", "callAPI_DoWork");

                Timer = new Timer(callback: new TimerCallback(startRealProcess), null, TimerLag, TimerPeriod);

            }
            catch (Exception ex)
            {
                PrintLog(ex.ToString().Remove(150), "", "", "OnStart"); // Log log error   
            }

        }

        protected override void OnStop()
        {
        }

        public void startRealProcess(object sender)
        {
            if (!callAPI.IsBusy)
                callAPI.RunWorkerAsync();
        }
        private void callAPI_DoWork(object sender, DoWorkEventArgs e)
        {
            //Api that calls a T24 Routine
            string apiUrl = ConfigurationManager.AppSettings["APIUrl"];
            string requestBody = "{\"routineName\":\"NAME.ROUTINE\",\"inParameters\":[\"data\",\"resp\"]}";
           
            
            try
            {

                string response = httpRequest(apiUrl, requestBody);

                if (response == "No File to Process")
                {
                   return;
                }
                else
                {
                    var logEntries = JsonConvert.DeserializeObject<List<Root>>(response);
                   
                    foreach (var entry in logEntries)
                    {
                        var datePart = entry.t_date;
                        var timePart = entry.t_time;
                        var DayPart = entry.t_type;
                        var ipAddress = entry.t_ip;

                        int midHrs = -1;

                         string tocc = ConfigurationManager.AppSettings["ITgroupmail"].Trim() + ", " +
                                 ConfigurationManager.AppSettings["RiskDeptmail"].Trim();

                       

                        if (DateTime.TryParseExact(datePart + " " + timePart, "dd MMM yyyy HH:mm:ss", null, System.Globalization.DateTimeStyles.None, out DateTime dateTime))
                        {
                            string weekday = dateTime.ToString("dddd");

                            midHrs = dateTime.Hour;

                            try
                            {

                                string to = entry.email_address;
                                string msg = entry.description_msg;
                                string MsgWithIP = entry.description_msg + " with IP " + ipAddress;
                                string phone = entry.phone_number;
                                string subj = entry.description_title;


                                if (midHrs < 7 || midHrs > 19 || weekday.Contains("Sunday") || entry.t_type == "HOLIDAY")
                                {
                                    msg = MsgWithIP;

                                    if (!string.IsNullOrEmpty(to) && ConfigurationManager.AppSettings["sendEmail"].Trim().Equals("true"))
                                        SendEmail("", "", subj, to, tocc, msg, "");

                                    if (!string.IsNullOrEmpty(phone))
                                    {
                                        SendSMS(phone, msg);
                                    }
                                    else
                                    {
                                        PrintLog("SMS not sent to Owner's Phone number as it is null or empty.", "1", "", "SendSMS");
                                    }

                                    if (new[] { "created", "debited pl" }.Any(word => entry.description_msg.ToLower().Contains(word)))
                                    {

                                        
                                            if (!string.IsNullOrEmpty(phone))
                                            {
                                                SendSMS(phone, msg);
                                            }
                                            else
                                            {
                                                PrintLog("SMS not sent to Owner's Phone number as it is null or empty.", "1", "", "SendSMS");
                                            }
                                        

                                        SendSMS(ConfigurationManager.AppSettings["ITPhone"].Trim(), msg);
                                        SendSMS(ConfigurationManager.AppSettings["RiskPhone"].Trim(), msg);
                                    }
                                }
                                else
                                {
                                    //if (entry.description_msg.ToLower().Contains("created") || entry.description_msg.ToLower().Contains("debited pl"))
                                    if (new[] { "created", "debited pl" }.Any(word => entry.description_msg.ToLower().Contains(word)))
                                    {

                                       if (!string.IsNullOrEmpty(to) && ConfigurationManager.AppSettings["sendEmail"].Trim().Equals("true"))
                                            SendEmail("", "", subj, to, tocc, msg, "");

                                    }
                                    else
                                    {
                                        if (!string.IsNullOrEmpty(to) && ConfigurationManager.AppSettings["sendEmail"].Trim().Equals("true"))
                                            SendEmail("", "", subj, to, "", msg, "");
                                       
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                PrintLog(ex.Message, "1", "", "callAPI_DoWork");
                            }
                        }
                        else
                        {
                            PrintLog("Error parsing the date.", "1", "", "callAPI_DoWork1");
                        }


                       string SendSMS(string number, string msg)
                        {
                            try
                            {
                                string str = "";
                                string url = ConfigurationManager.AppSettings["SMSURL"].Trim();

                                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls
                                       | SecurityProtocolType.Tls11
                                       | SecurityProtocolType.Tls12
                                       | SecurityProtocolType.Ssl3;

                                ServicePointManager.ServerCertificateValidationCallback =
                                delegate (object s, X509Certificate certificate, X509Chain chain, SslPolicyErrors sslPolicyErrors) { return true; };

                                HttpWebRequest httpWebRequest = WebRequest.Create(url) as HttpWebRequest;
                                httpWebRequest.Method = "POST";
                                httpWebRequest.Accept = "*/*";
                                httpWebRequest.ContentType = "application/xml";

                                string requestbody = "<sms id = \"D20220929144884722500.1\">" +
                                                      "<sender>25677.....</sender>" +
                                                      "<receiver>" + number + "</receiver>" +
                                                      "<message> " + msg + " </message>" +
                                                     "</sms>";

                                byte[] buffer = Encoding.UTF8.GetBytes(requestbody);
                                httpWebRequest.ContentLength = (long)buffer.Length;
                                httpWebRequest.ServicePoint.Expect100Continue = true;

                                using (Stream requestStream = httpWebRequest.GetRequestStream())
                                    requestStream.Write(buffer, 0, buffer.Length);

                                using (HttpWebResponse response1 = httpWebRequest.GetResponse() as HttpWebResponse)
                                    str = new StreamReader(response1.GetResponseStream()).ReadToEnd();
                                
                                return str;
                            }
                            catch (Exception ex)
                            {
                                PrintLog(ex.Message, "1", "", "SendSMS1");

                                return ex.Message;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
             
                string errorMessage = "An error occurred while calling the SendSMS2: ";

                if (ex != null)
                {
                    errorMessage += ex.Message;

                    if (ex.InnerException != null)
                    {
                        errorMessage += " Inner Exception: " + ex.InnerException.Message;
                    }

                    errorMessage += Environment.NewLine + "Stack Trace: " + ex.StackTrace;

                    // Log the full response
                    PrintLog(errorMessage, "1", "", "SendSMS2");
                }
                else
                {
                    // If ex is null, log a default message
                    PrintLog("An exception occurred in SendSMS2 but the error details are null.", "1", "", "SendSMS2");
                }
            
            }
        }

        static void SendEmail(string folder, string file_Name, string subject, string to, string ccTo, string Message, string disp)
        {
            try
            {
                string str2 = "";
                Message += "\n\n\n";
                string mailBox = ConfigurationManager.AppSettings["mailbox"].Trim();
                string emailPassword = ConfigurationManager.AppSettings["mailpass"].Trim();
                string serv_ip = ConfigurationManager.AppSettings["mailIP"].Trim();
                int port_Numebr = int.Parse(string.IsNullOrEmpty(ConfigurationManager.AppSettings["mailport"].Trim()) ? "5.." :
                    ConfigurationManager.AppSettings["mailport"].Trim());
                string from = ConfigurationManager.AppSettings["sender"].Trim();
                string displayName = ConfigurationManager.AppSettings["senderName"].Trim();
                subject = string.IsNullOrEmpty(subject) ? ConfigurationManager.AppSettings["subject"].Trim() : subject;
                displayName = string.IsNullOrEmpty(disp) ? displayName : disp;


                SmtpClient client = new SmtpClient(serv_ip, port_Numebr)
                {
                    Credentials = new NetworkCredential(mailBox, emailPassword)
                };
                MailMessage message = new MailMessage
                {
                    From = new MailAddress(from, displayName)
                };

                char ch = ',';
                if (ccTo.Trim() != string.Empty)
                {
                    string[] strArray = ccTo.Split(new char[] { ch });
                    if (ccTo.Contains<char>(ch))
                    {
                        foreach (string str3 in strArray)
                        {
                            message.CC.Add(new MailAddress(str3));
                        }
                    }
                    else
                    {
                        message.CC.Add(new MailAddress(ccTo));
                    }
                }

                if (to.Trim() != string.Empty)
                {
                    string[] strArray2 = to.Split(new char[] { ch });
                    if (to.Contains<char>(ch))
                    {
                        foreach (string str4 in strArray2)
                        {
                            message.To.Add(new MailAddress(str4));
                        }
                    }
                    else
                    {
                        message.To.Add(new MailAddress(to));
                    }
                }

                message.Subject = subject + str2;
                message.Priority = MailPriority.High;
                message.Body = Message;
                message.DeliveryNotificationOptions = DeliveryNotificationOptions.OnSuccess;

                MailAddress sender1 = new MailAddress(from, disp);
                message.Sender = sender1;

                if (string.IsNullOrEmpty(file_Name) == false || string.IsNullOrEmpty(folder) == false)
                {
                    if (string.IsNullOrEmpty(file_Name) == false)
                    {
                        message.Attachments.Add(new Attachment(folder + @"\" + file_Name));
                    }
                }

                client.Send(message);
                PrintLog("Mail sent " + to + " " + subject, "2", Message, "SendEmail");
                message.Dispose();

            }
            catch (Exception ex)
            {
                PrintLog(ex.Message, "1", "", "SendEmail1");
            }
        }

        static void PrintLog(string msg, string reg, string usr, string prFunc)
        {
            try
            {
                string tDay = DateTime.Now.ToString("yyyyMMdd");
                string dayTime = DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss") + " --:";
                string fLocal = ConfigurationManager.AppSettings["logfile123"].Trim();
                string file = fLocal + @"//" + tDay + ".txt";
                msg = dayTime + " - " + prFunc + ":- " + msg + "" + usr + " " + "\r\n";

                if (!Directory.Exists(fLocal))
                {
                    DirectoryInfo di = Directory.CreateDirectory(fLocal); // Try to create the directory.
                }

                switch (reg)
                {
                    case "0":
                        File.AppendAllText(file, msg);
                        break;
                    case "1":
                        File.AppendAllText(fLocal + "//Catcherr-logs-" + tDay + ".txt", msg);
                        break;
                    case "2":
                        File.AppendAllText(fLocal + "//Notification-logs-" + tDay + ".txt", msg);
                        break;
                }

            }
            catch (Exception ex)
            {
                string errormsg2 = "printLog" + ex.ToString().Remove(130);
            }
        }


        public static string httpRequest(string url, string data)
        {
            string str = string.Empty;
            ServicePointManager.Expect100Continue = true;
            ServicePointManager.ServerCertificateValidationCallback = delegate { return true; };
            try
            {
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
            }
            catch (Exception ex)
            {
                PrintLog(ex.Message, "1", "", "SendEmail");
            }

            HttpWebRequest httpWebRequest = WebRequest.Create(url) as HttpWebRequest;
            httpWebRequest.Method = "POST";
            httpWebRequest.ContentType = "application/json";
            byte[] buffer = Encoding.UTF8.GetBytes(data);
            httpWebRequest.ContentLength = buffer.Length;
            httpWebRequest.ServicePoint.Expect100Continue = false;

            try
            {
                using (Stream requestStream = httpWebRequest.GetRequestStream())
                    requestStream.Write(buffer, 0, buffer.Length);

                try
                {
                    using (HttpWebResponse response = httpWebRequest.GetResponse() as HttpWebResponse)
                    {
                        using (StreamReader reader = new StreamReader(response.GetResponseStream()))
                            str = reader.ReadToEnd();
                    }
                }
                catch (WebException ex)
                {
                    str = new StreamReader(ex.Response.GetResponseStream()).ReadToEnd();
                    PrintLog(str, "1", "", "SendEmail2");
                }
            }
            catch (Exception ex)
            {
                PrintLog(ex.Message, "1", "", "SendEmail3");
            }

            return str;
        }


        public class Root
        {
            public string phone_number { get; set; }
            public string email_address { get; set; }
            public string description_msg { get; set; }
            public string description_title { get; set; }
            public string t_type { get; set; }
            public string t_time { get; set; }
            public string t_date { get; set; }
            public string t_ip { get; set; }

        }
    }


}
