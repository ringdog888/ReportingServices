using System;
using System.Collections.Generic;
using System.IO;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
using System.Configuration;
using DocumentFormat.OpenXml.Bibliography;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Presentation;
using NPOI.SS.UserModel;
using static System.Runtime.CompilerServices.RuntimeHelpers;
using DocumentFormat.OpenXml.Vml.Spreadsheet;



namespace ReportingServices
{
    public partial class CTBC_02_8_Engagement_Plan : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            RptTool RptTool = new RptTool();
            //string EinB64 = Request.QueryString["EinB64"];
            //byte[] b = Convert.FromBase64String(EinB64);
            //string bStr = System.Text.Encoding.UTF8.GetString(b);
            //string[] baseStr = bStr.Split('&');
            string audit_no = "";
            string lang = "";


            //string connectionString = RptTool.LoadCmdStr("\\\\Database\\\\Project\\\\BPM\\\\BPMPro\\\\Connection\\\\Audit.xdbc.xmf");
            string connectionString = "Password=New@type;Persist Security Info=True;User ID=sa;Initial Catalog=Audit;Data Source=192.168.7.110;";

            //string connectionString = "Password=Astern@123;Persist Security Info=True;User ID=sa;Initial Catalog=Audit;Data Source=127.0.0.1;";

            string NUP = ConfigurationManager.AppSettings["NUP"];
            //Base64解碼
            /*foreach (string qstr in baseStr)
            {
                string[] qs = qstr.Split('=');
                if (qs[0] == "audit_no") audit_no = qs[1];
                if (qs[0] == "lang") lang = qs[1];
            }*/
            
            audit_no = "112-017";
            lang = "zh-tw";
            Dictionary<string, string> dic = new Dictionary<string, string>();
            dic.Add("@audit_no", audit_no);
            dic.Add("@lang", lang);

            //第二版--START
            System.Data.DataTable dt = RptTool.ExecSqlQueryParameters(connectionString, @"SELECT   a_list.[startdate] AS startdate,  
            a_list.[enddate] AS enddate,   DATEADD(YEAR, 5, enddate) AS Ptitleenddate,  [dbo].[Auditfn_GetLang_plantype_list](plantypeid,'zh-tw') AS plantype,
            a_list.audit_no AS auditno,   a_list.[searchTransactionName] AS planname ,
            a_assign.audit_range_startdate as ar_startdate, a_assign.audit_range_enddate as ar_enddate,
            (select [dbo].[Orgfn_GetLangDeptName](DeptID,'zh-tw')+ '、' from [dbo].[audit_system_auditplan_depts] 
            where planid=a_list.guidid) AS auditplandept,  (select [dbo].[Orgfn_GetLangMemName](accountid,'zh-tw') 
            from audit_system_auditplan_members where planid=a_list.guidid and isleader=1) AS leader,
            SUBSTRING( (select ([dbo].[Orgfn_GetLangMemName](accountid,'zh-tw'))+'、' from audit_system_auditplan_members 
            where planid=a_list.guidid and isleader=0 for xml path(''))  , 1,  
            LEN((select ([dbo].[Orgfn_GetLangMemName](accountid,'zh-tw'))+'、' from audit_system_auditplan_members 
            where planid=a_list.guidid and isleader=0 for xml path(''))) - 1)   AS Member        
            FROM [dbo].[audit_system_auditplan_list] AS a_list    
            JOIN [dbo].[audit_system_auditplan_assignments] as a_assign ON a_list.guidid = a_assign.planid
            JOIN [dbo].[audit_system_plantype_list] AS p_list  
            ON a_list.[plantypeid] = p_list.[guidid] where audit_no='112-017' ", dic);
            //第二版--END

            //    正確版本--START 但科目代碼無法帶出
            //            System.Data.DataTable dt = RptTool.ExecSqlQueryParameters(connectionString, @" SELECT   a_list.[startdate] AS startdate,  
            //            a_list.[enddate] AS enddate, DATEADD(YEAR, 5, enddate) AS Ptitleenddate,  [dbo].[Auditfn_GetLang_plantype_list] (plantypeid, 'zh-tw') AS plantype,
            //            a_list.audit_no AS auditno,   a_list.[searchTransactionName] AS planname,
            //            a_assign.audit_range_startdate as ar_startdate, a_assign.audit_range_enddate as ar_enddate,　
            //            a_assign.working_paperid,dbo.Auditfn_GetLang_workingpaper_list(a_assign.working_paperid, 'zh-tw') AS L1Name,

            //(select[dbo].[Orgfn_GetLangDeptName](DeptID, 'zh-tw') + '、' from[dbo].[audit_system_auditplan_depts]
            //where planid = a_list.guidid) AS auditplandept,

            //(select[dbo].[Orgfn_GetLangMemName](accountid, 'zh-tw')
            //from audit_system_auditplan_members where planid = a_list.guidid and isleader = 1) AS leader,
            //SUBSTRING( (select([dbo].[Orgfn_GetLangMemName](accountid, 'zh-tw')) + '、' from audit_system_auditplan_members
            //where planid = a_list.guidid and isleader = 0 for xml path(''))  , 1,  
            //LEN((select([dbo].[Orgfn_GetLangMemName](accountid, 'zh-tw')) + '、' from audit_system_auditplan_members
            //where planid = a_list.guidid and isleader = 0 for xml path(''))) -1)   AS Member,
            //(select REPLACE(Code,' ','') from [" + NUP + @"].[dbo].[CTBC_Company] where CompanyID=(select CompanyID from FSe7en_Org_DeptStruct where DeptID=(
            //select field_value from audit_system_auditplan_customfields where planid=a_list.guidid and field_name =
            //(select TOP(1)guidid from audit_system_plantype_custom_field where plantypeid=a_list.plantypeid and field_name='所屬科別')
            //))) AS CompanyName, 
            //(select spi.thecode from[dbo].[audit_system_parameter_item] as spi
            //join[dbo].[audit_system_workingpaper_businesscode] as wb on spi.guidid = wb.business_code
            //join[dbo].[audit_system_auditplan_assignments] as a_assign on a_assign.working_paperid = wb.business_code
            //where a_list.guidid = a_assign.planid) as subcode

            //FROM[dbo].[audit_system_auditplan_list] AS a_list
            //left join audit_system_auditplan_assignments as a_assign
            //on a_list.guidid = a_assign.planid
            //left join audit_system_auditplan_workingpaper_list wlist
            //on wlist.ori_guid = a_assign.working_paperid and wlist.planid = a_assign.planid
            //left join audit_system_auditplan_workingpaper_item as item
            //on item.working_paperid = wlist.guidid
            //JOIN[dbo].[audit_system_plantype_list] AS p_list
            //ON a_list.[plantypeid] = p_list.[guidid] where audit_no = '2023-017'", dic);
            //正確版本到這邊--END

            //最初寫法--START
            //DataTable dt = RptTool.ExecSqlQueryParameters(connectionString,
            //                 @"(SELECT a_list.[startdate] AS startdate, a_list.[enddate] AS enddate, DATEADD(YEAR, 5, enddate) AS Ptitleenddate, p_list.[guidid] AS plantypeid,
            //                 a_list.audit_no AS auditno, a_list.[searchTransactionName] AS planname FROM [dbo].[audit_system_auditplan_list] AS a_list JOIN [dbo].[audit_system_plantype_list] AS p_list 
            //                 ON a_list.[plantypeid] = p_list.[guidid],                         

            //                 (select [dbo].[Orgfn_GetLangMemName](accountid,@lang) from audit_system_auditplan_members where planid=list.guidid and isleader=1) AS leader,

            //                 SUBSTRING(
            //                 (select ([dbo].[Orgfn_GetLangMemName](accountid,@lang))+'、' from audit_system_auditplan_members where planid=list.guidid and isleader=0 for xml path(''))
            //                 , 1, LEN((select ([dbo].[Orgfn_GetLangMemName](accountid,@lang))+'、' from audit_system_auditplan_members where planid=list.guidid and isleader=0 for xml path(''))) - 1)
            //                 AS Member,
            //                (select REPLACE(Code,' ','') from [" + NUP + @"].[dbo].[CTBC_Company] where CompanyID=(select CompanyID from FSe7en_Org_DeptStruct where DeptID=(
            //                select field_value from audit_system_auditplan_customfields where planid=list.guidid and field_name =
            //                (select TOP(1)guidid from audit_system_plantype_custom_field where plantypeid=list.plantypeid and field_name='所屬科別')
            //                ))) AS CompanyName,    

            //                (select p_list.defaultname from [dbo].[audit_system_plantype_list] as p_list inner join 
            //                [dbo].[audit_system_auditplan_list] as a_list on a_list.plantypeid = p_list.guidid) AS plantype,

            //                (select [dbo].[Orgfn_GetLangDeptName](DeptID,@lang)+ '、' from [dbo].[audit_system_auditplan_depts]
            //                where planid=list.guidid) AS auditplandept", dic);
            //最初寫法--END

            if (dt.Rows.Count > 0)
            {
                if (lang == "zh-tw")
                {
                    CTBC_02_08_OPENXML.GeneratedClass Rpt = new CTBC_02_08_OPENXML.GeneratedClass();
                    string tempFile = Server.MapPath(".") + Guid.NewGuid() + ".docx";
                    Rpt.dt = dt;
                    Rpt.CreatePackage(tempFile);
                    byte[] buff = File.ReadAllBytes(tempFile);
                    Response.Clear();
                    Response.ContentType = "application/octet-stream";
                    Response.AddHeader("content-disposition", "attachment;filename=" + "內部稽核查核計畫.docx");
                    Response.Charset = "utf-8";
                    Response.BinaryWrite(buff);
                    File.Delete(tempFile);
                }
                else
                {
                    CTBC_02_08_OPENXML.GeneratedClass_en Rpt = new CTBC_02_08_OPENXML.GeneratedClass_en();
                    string tempFile = Server.MapPath(".") + Guid.NewGuid() + ".docx";
                    Rpt.dt = dt;
                    Rpt.CreatePackage(tempFile);
                    byte[] buff = File.ReadAllBytes(tempFile);
                    Response.Clear();
                    Response.ContentType = "application/octet-stream";
                    Response.AddHeader("content-disposition", "attachment;filename=" + "Engagement Plan.docx");
                    Response.Charset = "utf-8";
                    Response.BinaryWrite(buff);
                    File.Delete(tempFile);
                }
            }

            else
            {
                Response.Write("No Data");
            }
            Response.End();
        }
    }
}