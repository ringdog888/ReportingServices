using System;
using System.Collections.Generic;
using System.IO;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
using System.Data.SqlClient;
using System.Data.OleDb;
using Microsoft.Win32;
using DocumentFormat.OpenXml.Bibliography;

namespace ReportingServices
{
    public partial class CTBC_CheckCard : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            RptTool RptTool = new RptTool();
            string EinB64 = Request.QueryString["EinB64"] ?? "YXVkaXRfbm89MTEyLTAyMiZsYW5nPXpoLXR3Jk1haWxUYWc9MQ==";
            byte[] b = Convert.FromBase64String(EinB64);
            string bStr = System.Text.Encoding.UTF8.GetString(b);
            string[] baseStr = bStr.Split('&');
            string audit_no = "";
            string lang = "";
            string MailTag = "";
            string connectionString = RptTool.LoadCmdStr("\\\\Database\\\\Project\\\\BPM\\\\BPMPro\\\\Connection\\\\Audit.xdbc.xmf");
            //string connectionString = "Password=New@type;Persist Security Info=True;User ID=sa;Initial Catalog=Audit;Data Source=192.168.7.110;";


            //Base64解碼
            foreach (string qstr in baseStr)
            {
                string[] qs = qstr.Split('='); 
                if (qs[0] == "audit_no") audit_no = qs[1];
                if (qs[0] == "lang") lang = qs[1];
                if (qs[0] == "MailTag") MailTag = qs[1];
            }
            Dictionary<string, string> dic = new Dictionary<string, string>();
            dic.Add("@audit_no", audit_no);
            dic.Add("@lang", lang);
            DataTable dt = RptTool.ExecSqlQueryParameters(connectionString,
                    @"select CAST(GETDATE() AS DATE) AS printdate,list.guidid,startdate,enddate,audit_no,
                        [dbo].[Auditfn_GetLang_plantype_list](plantypeid,@lang) AS plantype,
                        (
                        select field_value from audit_system_auditplan_customfields where planid=list.guidid and field_name =
                        (select TOP(1)guidid from audit_system_plantype_custom_field where plantypeid=list.plantypeid and field_name='查程名稱')
                        ) AS planname,
                        [dbo].[Auditfn_GetLang_Level5]((
                        select field_value from audit_system_auditplan_customfields where planid=list.guidid and field_name =
                        (select TOP(1)guidid from audit_system_plantype_custom_field where plantypeid=list.plantypeid and field_name='所屬科別')
                        ),@lang) AS Belong,
                        (select REPLACE(Code,' ','') from [NUP].[dbo].[CTBC_Company] where CompanyID=(select CompanyID from FSe7en_Org_DeptStruct where DeptID=(
                        select field_value from audit_system_auditplan_customfields where planid=list.guidid and field_name =
                        (select TOP(1)guidid from audit_system_plantype_custom_field where plantypeid=list.plantypeid and field_name='所屬科別')
                        ))) AS CompanyName,
                        (select [dbo].[Orgfn_GetLangMemName](accountid,@lang) from audit_system_auditplan_members where planid=list.guidid and isleader=1) AS leader,
                        (select EMail from FSe7en_Org_MemberInfo where AccountID=(select accountid from audit_system_auditplan_members where planid=list.guidid and isleader=1)) AS leaderMail,
                        SUBSTRING(
                        (select ([dbo].[Orgfn_GetLangMemName](accountid,@lang))+',' from audit_system_auditplan_members where planid=list.guidid and isleader=0 for xml path(''))
                        , 1, LEN((select ([dbo].[Orgfn_GetLangMemName](accountid,@lang))+',' from audit_system_auditplan_members where planid=list.guidid and isleader=0 for xml path(''))) - 1)
                        AS Member,
                        (select the_year from audit_system_auditplan_audit_no_list where audit_no=list.audit_no) as audit_year
                        from audit_system_auditplan_list as list
                        where audit_no=@audit_no
                        ", dic);
            if (dt.Rows.Count > 0)
            {
                if (MailTag == "1")
                {
                    if (lang == "zh-tw")
                    {
                        CheckCard_CTBC.GeneratedClass Rpt = new CheckCard_CTBC.GeneratedClass();
                        string tempFile = Server.MapPath(".") + Guid.NewGuid() + ".docx";
                        Rpt.dt = dt;
                        Rpt.CreatePackage(tempFile);
                        CheckCard_CTBC.SnedMailClass mail = new CheckCard_CTBC.SnedMailClass();
                        mail.dt = dt;
                        mail.lang = lang;
                        Response.Write(mail.SendMail(tempFile));
                        File.Delete(tempFile);
                    }
                    else
                    {
                        CheckCard_CTBC.GeneratedClass_en Rpt = new CheckCard_CTBC.GeneratedClass_en();
                        string tempFile = Server.MapPath(".") + Guid.NewGuid() + ".docx";
                        Rpt.dt = dt;
                        Rpt.CreatePackage(tempFile);
                        CheckCard_CTBC.SnedMailClass mail = new CheckCard_CTBC.SnedMailClass();
                        mail.dt = dt;
                        mail.lang = lang;
                        Response.Write(mail.SendMail(tempFile));
                        File.Delete(tempFile);
                    }
                }
                else
                {
                    if (lang == "zh-tw")
                    {
                        CheckCard_CTBC.GeneratedClass Rpt = new CheckCard_CTBC.GeneratedClass();
                        string tempFile = Server.MapPath(".") + Guid.NewGuid() + ".docx";
                        Rpt.dt = dt;
                        Rpt.CreatePackage(tempFile);
                        byte[] buff = File.ReadAllBytes(tempFile);
                        Response.Clear();
                        Response.ContentType = "application/octet-stream";
                        Response.AddHeader("content-disposition", "attachment;filename=" + dt.Rows[0]["audit_no"].ToString() + "_檢查證.docx");
                        Response.Charset = "utf-8";
                        Response.BinaryWrite(buff);
                        File.Delete(tempFile);
                    }
                    else
                    {
                        CheckCard_CTBC.GeneratedClass_en Rpt = new CheckCard_CTBC.GeneratedClass_en();
                        string tempFile = Server.MapPath(".") + Guid.NewGuid() + ".docx";
                        Rpt.dt = dt;
                        Rpt.CreatePackage(tempFile);
                        byte[] buff = File.ReadAllBytes(tempFile);
                        Response.Clear();
                        Response.ContentType = "application/octet-stream";
                        Response.AddHeader("content-disposition", "attachment;filename=" + dt.Rows[0]["audit_no"].ToString() + "_Audit Notification.docx");
                        Response.Charset = "utf-8";
                        Response.BinaryWrite(buff);
                        File.Delete(tempFile);
                    }
                }
            }
            else {
                Response.Write("No Data");
            }
            Response.End();
        }
    }
}