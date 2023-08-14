
using System.Data;
using System;
using System.Net;
using System.Net.Mail;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml;
using System.Collections.Generic;
using System.Configuration;

namespace CheckCard_CTBC
{
    public class SnedMailClass
    {
        // Data Source
        public DataTable dt { get; set; }
        public string lang { get; set; }
        public string SendMail(string filePath)
        {
            string Content = "";
            string subject = "";
            string fileName = "";
            if (lang == "zh-tw") {
                fileName = dt.Rows[0]["audit_no"].ToString() + "_檢查證.docx";
                subject = @"內部稽核檢查通知：["+ dt.Rows[0]["audit_no"].ToString() + @"]["+
                    dt.Rows[0]["planname"].ToString() + @"](["+ dt.Rows[0]["plantype"].ToString() + @"])";
                Content = @"單位主管您好, <BR>
茲依據「金融控股公司及銀行業內部控制及稽核制度實施辦法」，前往貴單位辦理檢查，即請查照並惠予協助為荷<BR>
查程資訊：<BR>
查程名稱：" + dt.Rows[0]["CompanyName"].ToString() + @"_ " + dt.Rows[0]["planname"].ToString() + @" (" + dt.Rows[0]["plantype"].ToString() + @")<BR>
查核期間：" + Convert.ToDateTime(dt.Rows[0]["startdate"]).ToString("yyyy-MM-dd") + @" ~ " + Convert.ToDateTime(dt.Rows[0]["enddate"]).ToString("yyyy-MM-dd") + @"<BR>
查核成員：領隊稽核：" + dt.Rows[0]["leader"].ToString() + @"<BR>
稽核：";
                string[] MemberArr = dt.Rows[0]["Member"].ToString().Split(',');
                foreach (string Member in MemberArr)
                {
                    Content += Member+"<BR>";
                }
                    Content +=@"謝謝您的協助，如有任何問題請不吝與我們聯繫";
            }
            else {
                fileName = dt.Rows[0]["audit_no"].ToString() + "_Audit Notification.docx";
                subject = @"Notification of CTBC Bank Head Office Audit：[" + dt.Rows[0]["audit_no"].ToString() + @"][" +
                    dt.Rows[0]["planname"].ToString() + @"]([" + dt.Rows[0]["plantype"].ToString() + @"])";
                Content += @"Dear <font color='red'>XXXXX</font>,<BR>
                        <BR>
                        Please kindly be informed that CTBC International Audit Division will conduct [" + dt.Rows[0]["audit_year"].ToString() + @"] [" + dt.Rows[0]["planname"].ToString() + @"] [" + dt.Rows[0]["plantype"].ToString() + @"] <font color='red'>onsite/offsite</font> between [" + Convert.ToDateTime(dt.Rows[0]["startdate"]).ToString("yyyy-MM-dd") + @"] to [" + Convert.ToDateTime(dt.Rows[0]["enddate"]).ToString("yyyy-MM-dd") + @"] from Taipei.

                        Attached is the CTBC Bank Audit Notification for your information.

                        The Year [" + dt.Rows[0]["audit_year"].ToString() + @"] audit focus for [" + dt.Rows[0]["planname"].ToString() + @"] [" + dt.Rows[0]["plantype"].ToString() + @"] will include but not limited to:<BR>
                        <font color='red'>●	xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx</font><BR>
                        <font color='red'>●	xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx</font><BR>
                        <font color='red'>●	xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx</font><BR>
                        <font color='red'>●	xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx</font><BR>
<BR>
                        We will provide the list of documents requested to <font color='red'>XXXXXXXXXX</font> later to facilitate the document preparation, <font color='red'>which will be requested to be provided in separate lots in consideration of the limited manpower during lockdown</font>.<BR>
<BR>
                        Thank you for your generous support and your team’s kind assistance in this regard, particularly during the coronavirus outbreak.<BR>
<BR>
                        Sincerely,<BR>
                        <font color='red'>XXXXXXX</font>
                        ";
            }
            try
            {
                System.Net.Mail.MailMessage msg = new System.Net.Mail.MailMessage();
                if (ConfigurationManager.AppSettings["SMTP_Server"] == "smtp.gmail.com")
                {
                    msg.To.Add("andy.wang@newtype.com.tw");
                }
                else {
                    msg.To.Add(dt.Rows[0]["leaderMail"].ToString());
                }
                //msg.To.Add(dt.Rows[0]["leaderMail"].ToString());
                msg.From = new MailAddress(ConfigurationManager.AppSettings["SMTP_FromMail"], ConfigurationManager.AppSettings["SMTP_FromName"], System.Text.Encoding.UTF8);
                /* 上面3個參數分別是發件人地址（可以隨便寫），發件人姓名，編碼*/
                msg.Subject = subject;//郵件標題
                msg.SubjectEncoding = System.Text.Encoding.UTF8;//郵件標題編碼
                msg.Body = Content;

                msg.BodyEncoding = System.Text.Encoding.UTF8;//郵件內容編碼 

                msg.IsBodyHtml = true;//是否是HTML郵件 
                                      //msg.Priority = MailPriority.High;//郵件優先級 

                System.Net.Mail.Attachment attachment;
                attachment = new System.Net.Mail.Attachment(filePath);
                attachment.Name = fileName;
                attachment.NameEncoding = System.Text.Encoding.UTF8;
                msg.Attachments.Add(attachment);


                SmtpClient client = new SmtpClient();
                client.Credentials = new System.Net.NetworkCredential(ConfigurationManager.AppSettings["SMTP_Account"], ConfigurationManager.AppSettings["SMTP_Password"]); //這裡要填正確的帳號跟密碼
                client.Host = ConfigurationManager.AppSettings["SMTP_Server"] ; //設定smtp Server
                client.Port = Int32.Parse(ConfigurationManager.AppSettings["SMTP_Port"]); //設定Port
                client.EnableSsl = true; //gmail預設開啟驗證
                client.Send(msg); //寄出信件
                client.Dispose();
                msg.Dispose();
                return "success";
            }
            catch (Exception ex)
            {
                return ex.ToString();
            }
        }

    }
}