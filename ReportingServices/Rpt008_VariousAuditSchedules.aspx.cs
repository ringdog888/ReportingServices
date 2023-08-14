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

namespace ReportingServices
{
    public partial class Rpt008_VariousAuditSchedules : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            RptTool RptTool = new RptTool();
            string EinB64 = Request.QueryString["EinB64"] ?? "cGxhbnR5cGVpZD1DRkNEOTlEQS0wN0QyLTQzNjMtQjNFQi0zMzM2Qjk2QjlCMjAmTW9udGg9NiZHcm91cElEPTc4QTkxMUQyLTFDNEItNDA1NS1CMUJBLTUwOEMxQzUxMzE3RQ==";
            byte[] b = Convert.FromBase64String(EinB64);
            string bStr = System.Text.Encoding.UTF8.GetString(b);
            string[] baseStr = bStr.Split('&');
            string plantypeid = "";
            string GroupID = "";
            string Month = "";
            string connectionString = RptTool.LoadCmdStr("\\\\Database\\\\Project\\\\BPM\\\\BPMPro\\\\Connection\\\\Audit.xdbc.xmf");
            //string connectionString = "Password=sa;Persist Security Info=True;User ID=sa;Initial Catalog=Audit;Data Source=WIN-7AVURFP25PL;";


            //Base64解碼
            foreach (string qstr in baseStr)
            {
                string[] qs = qstr.Split('=');
                if (qs[0] == "plantypeid") plantypeid = qs[1];
                if (qs[0] == "GroupID") GroupID = qs[1];
                if (qs[0] == "Month") Month = qs[1];
            }
            Dictionary<string, string> dic = new Dictionary<string, string>();
            dic.Add("@plantypeid", plantypeid);
            dic.Add("@Month", Month);
            dic.Add("@GroupID", GroupID);
            DataTable dt = RptTool.ExecSqlQueryParameters(connectionString,
                    @"select (@Month+'月份各項業務查核預定表 '+dbo.Auditfn_GetLang_dept_subgroup_list(@GroupID,'zh-tw')+' '+dbo.Auditfn_GetLang_plantype_list(plantype.guidid,'zh-tw'))
                    AS Title ,dbo.Orgfn_GetLangTag(depts.deptid,'zh-tw') AS deptname,depts.deptid AS DeptID,list.guidid,
                    '自'+FORMAT(CAST(startdate AS DATE), 'M月d日') as StartDate,'到'+FORMAT(CAST(enddate AS DATE), 'M月d日') AS EndDate,DATEDIFF(DAY, startdate, enddate)+1 AS Days
                    ,(select dbo.Orgfn_GetLangMemName(accountid,'zh-tw')+' ' from audit_system_auditplan_members 
                    where audit_system_auditplan_members.planid=List.guidid for xml path('')) AS Checkers
                    from audit_system_auditplan_list list
                    left join audit_system_auditplan_depts depts
                    on depts.planid=List.guidid
                    left join [audit_system_plantype_list] plantype
                    on list.plantypeid=plantype.guidid
                    where (
                        (select responsible_type from audit_org_dept_subgroup_list where guidid=@GroupID)='all'
                         or 
                        (depts.deptid in (
                        select responsible_id from audit_org_dept_subgroup_responsible where groupid=@GroupID)
                        and (select responsible_type from audit_org_dept_subgroup_list where guidid=@GroupID)<>'all'
                        )) and plantypeid=@plantypeid 
                    and MONTH(startdate)<=@Month and MONTH(enddate)>=@Month", dic);
            Rpt008.GeneratedClass Rpt = new Rpt008.GeneratedClass();
            string tempFile = Server.MapPath(".") + Guid.NewGuid() + ".docx";
            Rpt.dt = dt;
            Rpt.CreatePackage(tempFile);
            byte[] buff = File.ReadAllBytes(tempFile);
            File.Delete(tempFile);
            Response.Clear();
            Response.ContentType = "application/octet-stream";
            Response.AddHeader("content-disposition", "attachment;filename=各項稽查預定表.docx");
            Response.Charset = "utf-8";
            Response.BinaryWrite(buff);
            Response.End();
        }
    }
}