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
    public partial class Rpt002_CheckCard_Stub : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            RptTool RptTool = new RptTool();
            string EinB64 = Request.QueryString["EinB64"] ?? "WWVhcj0yMDIzJk1vbnRoPTYmR3JvdXBJRD0wJkRlcHRJRD1CMTAzMDUwMDAw";
            byte[] b = Convert.FromBase64String(EinB64);
            string bStr = System.Text.Encoding.UTF8.GetString(b);
            string[] baseStr = bStr.Split('&');
            string YY = "";
            string MM = "";
            string AC = "";
            string DID = "";
            string connectionString = RptTool.LoadCmdStr("\\\\Database\\\\Project\\\\BPM\\\\BPMPro\\\\Connection\\\\Audit.xdbc.xmf");
            //string connectionString = "Password=sa;Persist Security Info=True;User ID=sa;Initial Catalog=Audit;Data Source=WIN-7AVURFP25PL;";

            //Base64解碼
            foreach (string qstr in baseStr)
            {
                string[] qs = qstr.Split('=');
                if (qs[0] == "Year") YY = qs[1];
                if (qs[0] == "Month") MM = qs[1];
                if (qs[0] == "plantypeid") AC = qs[1];
                if (qs[0] == "DeptID") DID = qs[1];
            }
            string sqlstr = @"select DATEDIFF(DAY,StartDate,EndDate)+1 as DIFFDAY,
                            CAST((DAY(EndDate)) as nvarchar) as EDAY, 
                            CAST((DAY(StartDate)) as nvarchar) as SDAY, 
                            CAST((MONTH(EndDate)) as nvarchar) as EMONTH, 
                            CAST((MONTH(StartDate)) as nvarchar) as SMONTH, 
                            CAST((YEAR(EndDate)-1911) as nvarchar) as EYEAR, 
                            CAST((YEAR(StartDate)-1911) as nvarchar) as SYEAR, 
                            CAST((YEAR(GETDATE())-1911) as nvarchar) as YEAR, 
                            CAST((MONTH(GETDATE())) as nvarchar) as MONTH, 
                            CAST((DAY(GETDATE())) as nvarchar) as DAY, 
                            audit_no as UID,
                            dbo.Auditfn_GetLang_plantype_list(plantype.guidid,'zh-tw') as Class,
                            CASE  
                            WHEN CHARINDEX(dbo.Auditfn_GetLang_plantype_list(plantype.guidid,'zh-tw'),'一般')>=0 THEN '一' 
                            WHEN CHARINDEX(dbo.Auditfn_GetLang_plantype_list(plantype.guidid,'zh-tw'),'專案')>=0 THEN '專' ELSE '覆' END AuditClassName,
                            (select dbo.Orgfn_GetLangMemName(accountid,'zh-tw')+',' from audit_system_auditplan_members 
                             where audit_system_auditplan_members.planid=List.guidid and isleader=0 for xml path('')) AS ppl,
                             StartDate AS RealStartDate,enddate AS RealEndDate,
                             (select dbo.Auditfn_GetLang_dept_subgroup_list(groupid,'zh-tw') from  audit_org_dept_subgroup_responsible
                            where responsible_id=depts.deptid for xml path('')) AS deptname,
                             (select groupid from  audit_org_dept_subgroup_responsible
                            where responsible_id=depts.deptid for xml path('')) AS DeptID,
                            depts.deptid AS AuditDeptID
                            , '' AS ParentDeptID,0 as holiday
                            from audit_system_auditplan_list list
                            left join audit_system_auditplan_depts depts
                            on depts.planid=List.guidid
                            left join audit_system_plantype_list plantype
                            on list.plantypeid=plantype.guidid 
                            WHERE Year(StartDate) = @Year AND Month(StartDate) = @Month 
                             AND depts.deptid = @DeptID ";

            

            //Response.Write(sqlstr);
            //Response.End();

            Dictionary<string, string> dic = new Dictionary<string, string>();
            dic.Add("@DeptID",DID);
            dic.Add("@Year", YY);
            dic.Add("@Month", MM);
            if(AC != "")
            {
                sqlstr = sqlstr + @"and list.plantypeid=@plantypeid";
                dic.Add("@plantypeid", AC);
            }
            DataTable dt = RptTool.ExecSqlQueryParameters(connectionString, sqlstr, dic);

            Rpt002_Stub.GeneratedClass Rpt = new Rpt002_Stub.GeneratedClass();
            string tempFile = Server.MapPath(".") + Guid.NewGuid() + ".docx";
            Rpt.dt = dt;
            Rpt.CreatePackage(tempFile);
            byte[] buff = File.ReadAllBytes(tempFile);
            File.Delete(tempFile);
            string fileName = "查核證(存根).docx";
            if (Request.Browser.Browser == "IE")
            {
                fileName = Server.UrlPathEncode(fileName);
            }
            Response.Clear();
            Response.ContentType = "application/octet-stream";
            Response.AddHeader("content-disposition", "attachment;filename=" + fileName);
            Response.Charset = "utf-8";
            Response.BinaryWrite(buff);
            Response.End();
        }
    }
}