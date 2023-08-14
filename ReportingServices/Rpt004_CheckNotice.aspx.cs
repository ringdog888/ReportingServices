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
using DocumentFormat.OpenXml.Presentation;

namespace ReportingServices
{
    public partial class Rpt004_CheckNotice : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            RptTool RptTool = new RptTool();
            string EinB64 = Request.QueryString["EinB64"] ?? "Z3VpZGlkPTAwNTI2QzEyLUI0QUQtNDRCOS04NjM0LUM3RkYwMzI5QTdCMCZEZXB0SUQ9QjA5NDA1MDAwMCZtQ2xhc3M9MA==";
            byte[] b = Convert.FromBase64String(EinB64);
            string bStr = System.Text.Encoding.UTF8.GetString(b);
            string[] baseStr = bStr.Split('&');
            string UID = "6";
            // string connectionString = RptTool.LoadCmdStr("\\\\Database\\\\Project\\\\BPM\\\\BPMPro\\\\Connection\\\\Audit.xdbc.xmf");
            //string connectionString = "Password=sa;Persist Security Info=True;User ID=sa;Initial Catalog=Audit;Data Source=WIN-7AVURFP25PL;";
            string connectionString = "Password=Astern@123;Persist Security Info=True;User ID=sa;Initial Catalog=Audit;Data Source=127.0.0.1;";
            //Base64解碼
            foreach (string qstr in baseStr)
            {
                string[] qs = qstr.Split('=');
                //if (qs[0] == "UniqueID") UID = qs[1];
            }

            Dictionary<string, string> dic = new Dictionary<string, string>();
            dic.Add("@UniqueID", UID);
            string sqlstr = @" SELECT MS.UniqueID, MS.Stage AS Stage, Month(MS.sDate) AS sm, Day(MS.sDate) AS sd, Month(MS.eDate) AS em,
                             Day(MS.eDate) AS ed, MS.AuditDept AS AuditDept, MS.Contact AS Contact, PM.DeptName AS DeptName, 
                             PM.Title AS Title, PM.FullName AS FullName, TI.ItemFull AS ItemFull, 
                             YEAR(GETDATE()) AS Year, Month(GETDATE()) AS Month, Day(GETDATE()) AS Day 
                             FROM eAudit_Training_MemStage AS MS INNER JOIN 
                             eAudit_Training_PlanMem AS PM ON MS.TPMID = PM.UniqueID INNER JOIN 
                            eAudit_Training_Item AS TI ON MS.ItemID = TI.UniqueID 
                             WHERE MS.UniqueID= @UniqueID ";

            DataTable dt = RptTool.ExecSqlQueryParameters(connectionString, sqlstr, dic);

            Rpt004CheckNotice.GeneratedClass Rpt = new Rpt004CheckNotice.GeneratedClass();
            string tempFile = Server.MapPath(".") + Guid.NewGuid() + ".docx";
            Rpt.dt = dt;
            Rpt.CreatePackage(tempFile);
            byte[] buff = File.ReadAllBytes(tempFile);
            File.Delete(tempFile);
            string fileName = "查核實習通知單.docx";
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