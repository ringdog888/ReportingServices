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
    public partial class Rpt001_IAUnterlagenOverseas : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {   
            RptTool RptTool = new RptTool();
            string EinB64 = Request.QueryString["EinB64"]?? "R3JvdXBJRD1ERkNBMjk2MS0xQUI4LTQ5NEQtOEFDMC1CQkZFNDk3MEM5MDYmcGxhbnR5cGVpZD1DRkNEOTlEQS0wN0QyLTQzNjMtQjNFQi0zMzM2Qjk2QjlCMjA=";
            byte[] b = Convert.FromBase64String(EinB64);
            string bStr = System.Text.Encoding.UTF8.GetString(b);
            string[] baseStr = bStr.Split('&');
            string GroupID = "";
            string plantypeid = "";
            string connectionString = RptTool.LoadCmdStr("\\\\Database\\\\Project\\\\BPM\\\\BPMPro\\\\Connection\\\\Audit.xdbc.xmf");
            //string connectionString = "Password=sa;Persist Security Info=True;User ID=sa;Initial Catalog=Audit;Data Source=WIN-7AVURFP25PL;";


            //Base64解碼
            foreach (string qstr in baseStr)
            {
                string[] qs = qstr.Split('=');
                if (qs[0] == "GroupID") GroupID = qs[1];
                if (qs[0] == "plantypeid") plantypeid = qs[1];
            }

            Dictionary<string, string> dic = new Dictionary<string, string>();
            dic.Add("@GroupID", GroupID);
            dic.Add("@plantypeid", plantypeid);
            DataTable dt = RptTool.ExecSqlQueryParameters(connectionString, @"
                                                    select @GroupID AS GID,@plantypeid AS PID,assign.working_paperid,dbo.Auditfn_GetLang_workingpaper_list(assign.working_paperid,'zh-tw') AS L1Name from audit_system_auditplan_list list 
                                                    left join audit_system_auditplan_assignments assign
                                                    on list.guidid=assign.planid
                                                    left join audit_system_auditplan_workingpaper_list wlist
                                                    on wlist.ori_guid=assign.working_paperid and wlist.planid=assign.planid
                                                    left join audit_system_auditplan_workingpaper_item item
                                                    on item.working_paperid=wlist.guidid
                                                    left join audit_system_auditplan_depts depts
                                                    on depts.planid=assign.planid
                                                    where list.plantypeid=@plantypeid and (
                                                                            (select responsible_type from audit_org_dept_subgroup_list where guidid='')='all'
                                                                             or 
                                                                            (depts.deptid in (
                                                                            select responsible_id from audit_org_dept_subgroup_responsible where groupid=@GroupID)
                                                                            and (select responsible_type from audit_org_dept_subgroup_list where guidid=@GroupID)<>'all'
                                                                            ))
                                                    group by assign.working_paperid
                                                    ", dic);
            
            Rpt001.GeneratedClass Rpt = new Rpt001.GeneratedClass();
            string tempFile = Server.MapPath(".") + Guid.NewGuid() + ".docx";
            Rpt.dt_L1 = dt;
            Rpt.connectionString = connectionString;
            Rpt.CreatePackage(tempFile);
            byte[] buff = File.ReadAllBytes(tempFile);
            File.Delete(tempFile);
            Response.Clear();
            Response.ContentType = "application/octet-stream";
            Response.AddHeader("content-disposition", "attachment;filename=工作底稿.docx");
            Response.Charset = "utf-8";
            Response.BinaryWrite(buff);
            Response.End();
        }
    }
}