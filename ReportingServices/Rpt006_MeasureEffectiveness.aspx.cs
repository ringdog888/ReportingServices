using System;
using System.IO;
using System.Data;
using System.Collections.Generic;

namespace ReportingServices
{
    public partial class Rpt006_MeasureEffectiveness : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            RptTool RptTool = new RptTool();

            string EinB64 = Request.QueryString["EinB64"] ?? "WWVhcj0yMDEzJnNNb250aD0xJmVNb250aD0xMg==";
            byte[] b = Convert.FromBase64String(EinB64);
            string bStr = System.Text.Encoding.UTF8.GetString(b);
            string[] baseStr = bStr.Split('&');
            string RY = "";
            string SM = "";
            string EM = "";
            string connectionString = RptTool.LoadCmdStr("\\\\Database\\\\Project\\\\BPM\\\\BPMPro\\\\Connection\\\\Audit.xdbc.xmf");
            //string connectionString = "Password=sa;Persist Security Info=True;User ID=sa;Initial Catalog=Audit;Data Source=WIN-7AVURFP25PL;";

            //Base64解碼
            foreach (string qstr in baseStr)
            {
                string[] qs = qstr.Split('=');
                if (qs[0] == "Year") RY = qs[1];
                if (qs[0] == "sMonth") SM = qs[1];
                if (qs[0] == "eMonth") EM = qs[1];
            }

            string sqlstr =
                @"SELECT	'內部稽核-' + 
						    CASE AD.ParentDeptID 
						    WHEN '001' THEN '國內營業單位'
						    WHEN '002' THEN '國外營業單位'
						    WHEN '000' THEN '總行單位' 
						    ELSE PAD.DeptName END as [細項名稱], 
						    cast(YEAR(PL.RealStartDate) - 1911 as nvarchar) + 'Q' + cast((MONTH(PL.RealStartDate) - 1) / 3 + 1 as nvarchar) as [風險評估期別],
						    PL.DeptID as [受查主體代號],
						    PL.DeptName as [受查主體名稱],
						    RO.content as [風險類型],
						    AI.Summary as [評估情形],
						    case 
							    when CHARINDEX('應予糾正', GT.content) > 0 OR
									    CHARINDEX('重複未改善', GT.content) > 0 OR
									    CHARINDEX('防制洗錢', RO.content) > 0 OR
									    CHARINDEX('法令遵循', RO.content) > 0
							    then '高'
							    when CHARINDEX('一般作業缺失', GT.content) > 0
							    then '中' else '' end as [風險等級]
				    FROM	eAudit_IA_AuditItem AS AI INNER JOIN 
						    eAudit_IA_PlansMem AS PM ON AI.PMID = PM.UniqueID INNER JOIN 
						    eAudit_IA_Plans AS PL ON PM.PlanID = PL.UniqueID LEFT OUTER JOIN 
						    eAudit_AuditBase_LEVEL02 AS L2 ON AI.L2ID = L2.UniqueID LEFT OUTER JOIN 
						    eAudit_AuditBase_LEVEL01 AS L1 ON L2.L1ID = L1.UniqueID LEFT OUTER JOIN 
						    eAudit_AuditBase_GlossaryType AS GT ON AI.Type = GT.UniqueID LEFT OUTER JOIN 
						    AFS_Dept AS AD ON PL.DeptID = AD.DeptID LEFT OUTER JOIN 
						    AFS_Dept AS PAD ON AD.ParentDeptID = PAD.DeptID LEFT OUTER JOIN 
						    eAudit_IA_RiskOpinion AS RO ON AI.UniqueID = RO.AIID 
				    WHERE	(NOT (AI.Result = 1)) AND (PL.Event = 100) AND 
						    (YEAR(PL.RealStartDate) = '" + RY + @"') AND  
						    (MONTH(PL.RealStartDate) >= '" + SM + @"') AND 
						    (MONTH(PL.RealStartDate) <= '" + EM + @"')";

            Dictionary<string, string> dic = new Dictionary<string, string>();
            DataTable dt = RptTool.ExecSqlQueryParameters(connectionString, sqlstr, dic);

            Rpt006MeasureEffectiveness.GeneratedClass Rpt = new Rpt006MeasureEffectiveness.GeneratedClass();
            string tempFile = Server.MapPath(".") + Guid.NewGuid() + ".xlsx";
            Rpt.dt = dt;
            Rpt.CreatePackage(tempFile);
            byte[] buff = File.ReadAllBytes(tempFile);
            File.Delete(tempFile);
            string fileName = "第一銀行控制措施有效性評估表格.xlsx";
            fileName = Server.UrlPathEncode(fileName);
            Response.Clear();
            Response.ContentType = "application/octet-stream";
            Response.AddHeader("content-disposition", "attachment;filename=" + fileName);
            Response.Charset = "utf-8";
            Response.BinaryWrite(buff);
            Response.End();



        }
    }
}