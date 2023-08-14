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
    public partial class Rpt002_CheckCard_Notice : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            RptTool RptTool = new RptTool();
            string EinB64 = Request.QueryString["EinB64"]?? "WWVhcj0yMDIzJk1vbnRoPTYmR3JvdXBJRD0wJkRlcHRJRD1CMTAzMDUwMDAw";
            byte[] b = Convert.FromBase64String(EinB64);
            string bStr = System.Text.Encoding.UTF8.GetString(b);
            string[] baseStr = bStr.Split('&');
            string YY = "";
            string MM = "";	
            string AC = "";	
            string DID = "";
            //string connectionString = RptTool.LoadCmdStr("\\\\Database\\\\Project\\\\BPM\\\\BPMPro\\\\Connection\\\\Audit.xdbc.xmf");
            string connectionString = "Password=New@type;Persist Security Info=True;User ID=sa;Initial Catalog=Audit;Data Source=192.168.7.120;";

            //Base64解碼
            foreach (string qstr in baseStr)
            {
                string[] qs = qstr.Split('=');
				if (qs[0] == "Year") YY=qs[1];
				if (qs[0] == "Month") MM=qs[1];
				if (qs[0] == "AuditClass") AC=qs[1];
				if (qs[0] == "DeptID") DID=qs[1];
            }

            /*
            string sqlstr = " SELECT UniqueID, CAST((YEAR(GETDATE())-1911) as nvarchar) as YEAR, CAST((MONTH(GETDATE())) as nvarchar) as MONTH, CAST((DAY(GETDATE())) as nvarchar) as DAY, CAST((YEAR(GETDATE())-1911) as nvarchar)+RIGHT(REPLICATE('0', 3) + CAST(UniqueID as NVARCHAR), 3) as UID, CASE AuditClass WHEN 0 THEN '一' WHEN 1 THEN '專' ELSE '覆' END AuditClassName, AuditClass ,ProjectOption, (SELECT CASE AuditClass WHEN 0 THEN '一般查核' WHEN 4 THEN (select AccName from eAudit_Critical_PlanDetail as CP WHERE CP.PlanID = DATA.UniqueID ) WHEN 1 THEN '專案查核 - ' +" +
                    " ISNULL((SELECT OptionName FROM eAudit_AuditBase_Options WHERE (UniqueID = data.ProjectOption)) ,'內部管理') WHEN 9 THEN '應業務需要' ELSE name END " +
                    " FROM eAudit_AuditBase_AuditClass as AC WHERE (AuditClass =data.AuditClass)) as Class , DeptName, RealStartDate, RealEndDate, AuditDeptID, ISNULL(REPLACE(REPLACE(ppl, '', ''), '', '') ,'')+ISNULL(TrainingMem ,'') AS ppl, data.holiday FROM (SELECT CASE PL.TrainingMem WHEN NULL THEN '' WHEN '' THEN '' ELSE '實習查核員,'+TrainingMem +',' END as TrainingMem ,PL.UniqueID, PL.AuditClass, PL.DeptName, PL.RealStartDate, PL.RealEndDate, PL.AuditDeptID, ProjectOption, (SELECT AuditMemName + ',' FROM eAudit_IA_PlansMem WHERE (PlanID = PL.UniqueID) AND (NOT (AuditMemName IS NULL)) GROUP BY AuditMemName, Leader ORDER BY Leader DESC FOR XML PATH('')) AS ppl, (SELECT COUNT(*) AS HC FROM (SELECT EventDATE AS holiday FROM eAudit_Sys_Holiday " +
                            " WHERE (Exist = '1') AND (EventTAG = '0') AND (YEAR(EventDATE) = '" + YY + "') UNION " +
                            " SELECT NewDate FROM (SELECT CAST('" + YY + "' + '-' + CAST(MONTH(EventDATE) AS nvarchar) + '-' + CAST(DAY(EventDATE) AS nvarchar) AS DATE) " +
                            " AS NewDate FROM eAudit_Sys_Holiday AS eAudit_Sys_Holiday_1 " +
                            " WHERE (Exist = '1') AND (EventTAG = '1')) AS derivedtbl_1) AS derivedtbl_2 " +
                            " WHERE (holiday BETWEEN PL.RealStartDate AND PL.RealEndDate)) AS holiday " +
                            " FROM eAudit_IA_Plans AS PL) AS data " +
                            " WHERE not(AuditClass in(5,6,7)) " +
                            " AND Year(data.RealStartDate) = '" + YY + "' AND Month(data.RealStartDate) = '" + MM + "' " +
                            " AND AuditDeptID = '" + DID + "' ";
           

            string sqlstr = @"
            SELECT UniqueID,
                    CAST((YEAR(GETDATE())-1911) AS nvarchar) AS YEAR,
                    CAST((MONTH(GETDATE())) AS nvarchar) AS MONTH,
                    CAST((DAY(GETDATE())) AS nvarchar) AS DAY,
                    CAST((YEAR(GETDATE())-1911) AS nvarchar)+RIGHT(REPLICATE('0', 3) + CAST(UniqueID AS NVARCHAR), 3) AS UID,
                    CASE AuditClass
                        WHEN 0 THEN '一'
                        WHEN 1 THEN '專'
                        ELSE '覆'
                    END AuditClassName,
                    AuditClass,
                    ProjectOption,

                (SELECT CASE AuditClass
                            WHEN 0 THEN '一般查核'
                            WHEN 4 THEN
                                    (SELECT AccName
                                    FROM eAudit_Critical_PlanDetail AS CP
                                    WHERE CP.PlanID = DATA.UniqueID )
                            WHEN 1 THEN '專案查核 - ' + ISNULL(
                                                            (SELECT OptionName
                                                            FROM eAudit_AuditBase_Options
                                                            WHERE (UniqueID = data.ProjectOption)) ,'內部管理')
                            WHEN 9 THEN '應業務需要'
                            ELSE name
                        END
                FROM eAudit_AuditBase_AuditClass AS AC
                WHERE (AuditClass =data.AuditClass)) AS CLASS,
                    DeptName,
                    RealStartDate,
                    RealEndDate,
                    AuditDeptID,
                    ISNULL(REPLACE(REPLACE(ppl, '', ''), '', ''), '')+ISNULL(TrainingMem, '') AS ppl,
                    data.holiday
            FROM
                (SELECT CASE PL.TrainingMem
                            WHEN NULL THEN ''
                            WHEN '' THEN ''
                            ELSE '實習查核員,'+TrainingMem +','
                        END AS TrainingMem,
                        PL.UniqueID,
                        PL.AuditClass,
                        PL.DeptName,
                        PL.RealStartDate,
                        PL.RealEndDate,
                        PL.AuditDeptID,
                        ProjectOption,

                    (SELECT AuditMemName + ','
                    FROM eAudit_IA_PlansMem
                    WHERE (PlanID = PL.UniqueID)
                    AND (NOT (AuditMemName IS NULL))
                    GROUP BY AuditMemName,
                            Leader
                    ORDER BY Leader DESC
                    FOR XML PATH('')) AS ppl,
                    (SELECT COUNT(*) AS HC
                    FROM
                    (SELECT EventDATE AS holiday
                        FROM eAudit_Sys_Holiday
                        WHERE (Exist = '1')
                        AND (EventTAG = '0')
                        AND (YEAR(EventDATE) = '" + YY + @"')
                        UNION SELECT NewDate
                        FROM
                        (SELECT CASE 
	                            WHEN DAY(EventDATE) > day(dateadd(month,1,'" + YY + @"-'+CAST(MONTH(EventDATE) AS nvarchar)+'-01')-1)
	                            THEN NULL 
	                            ELSE 
		                            CAST('" + YY + @"' + '-' + CAST(MONTH(EventDATE) AS nvarchar) + '-' + CAST(DAY(EventDATE) AS nvarchar) AS DATE)
	                            END AS NewDate
                        FROM eAudit_Sys_Holiday AS eAudit_Sys_Holiday_1
                        WHERE (Exist = '1')
                            AND (EventTAG = '1')) AS derivedtbl_1) AS derivedtbl_2
                    WHERE (holiday BETWEEN PL.RealStartDate AND PL.RealEndDate)) AS holiday
                FROM eAudit_IA_Plans AS PL) AS DATA
            WHERE not(AuditClass in(5, 6, 7))
                AND Year(data.RealStartDate) = '" + YY + @"'
                AND Month(data.RealStartDate) = '" + MM + @"'
                AND AuditDeptID = '" + DID + @"' "; */
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
            dic.Add("@DeptID", DID);
            dic.Add("@Year", YY);
            dic.Add("@Month", MM);
            if (AC != "")
            {
                sqlstr = sqlstr + @"and list.plantypeid=@plantypeid";
                dic.Add("@plantypeid", AC);
            }
            DataTable dt = RptTool.ExecSqlQueryParameters(connectionString, sqlstr, dic);

            Rpt002_Notice.GeneratedClass Rpt = new Rpt002_Notice.GeneratedClass();
            string tempFile = Server.MapPath(".") + Guid.NewGuid() + ".docx";
            Rpt.dt = dt;
            Rpt.CreatePackage(tempFile);
            byte[] buff = File.ReadAllBytes(tempFile);
            File.Delete(tempFile);
            string fileName = "查核證(通知).docx";
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