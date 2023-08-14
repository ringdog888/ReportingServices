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
using DocumentFormat.OpenXml.Wordprocessing;

namespace ReportingServices
{
    public partial class Rpt003_PracticeCard_2 : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            RptTool RptTool = new RptTool();
            string EinB64 = Request.QueryString["EinB64"] ?? "Z3VpZGlkPTAwNTI2QzEyLUI0QUQtNDRCOS04NjM0LUM3RkYwMzI5QTdCMCZEZXB0SUQ9QjA5NDA1MDAwMCZtQ2xhc3M9MA==";
            byte[] b = Convert.FromBase64String(EinB64);
            string bStr = System.Text.Encoding.UTF8.GetString(b);
            string[] baseStr = bStr.Split('&');
            string ID = "";
            string connectionString = RptTool.LoadCmdStr("\\\\Database\\\\Project\\\\BPM\\\\BPMPro\\\\Connection\\\\Audit.xdbc.xmf");
            //string connectionString = "Password=sa;Persist Security Info=True;User ID=sa;Initial Catalog=Audit;Data Source=WIN-7AVURFP25PL;";

            //Base64解碼
            foreach (string qstr in baseStr)
            {
                string[] qs = qstr.Split('=');
                if (qs[0] == "ID") ID = qs[1];
            }

            string sqlstr = @" SELECT UniqueID, Issue, bookDate, 
                                    (SELECT TOP (1) FullName 
                                      FROM eAudit_Training_PlanMem 
                                      WHERE (TPID = PL.UniqueID) AND (Exist = 1)) AS ppl, 
                                  (SELECT COUNT(*) AS Expr1 
                                     FROM eAudit_Training_PlanMem AS eAudit_Training_PlanMem_1 
                                     WHERE (TPID = PL.UniqueID) AND (Exist = 1)) AS p_num, 
                                    YEAR(GETDATE()) AS Year, 
                                    Month(GETDATE()) AS Month, 
                                    Day(GETDATE()) AS Day, 
                                     Accordance AS context 
                             FROM eAudit_Training_Plans AS PL 
            WHERE (UniqueID = 11) ";

            Dictionary<string, string> dic = new Dictionary<string, string>();
            DataTable dt = RptTool.ExecSqlQueryParameters(connectionString, sqlstr, dic);

            Rpt003_PracticeCard2.GeneratedClass Rpt = new Rpt003_PracticeCard2.GeneratedClass();
            string tempFile = Server.MapPath(".") + Guid.NewGuid() + ".docx";
            Rpt.dt = dt;
            Rpt.CreatePackage(tempFile);
            byte[] buff = File.ReadAllBytes(tempFile);
            File.Delete(tempFile);
            string fileName = "實習證明書2.docx";
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