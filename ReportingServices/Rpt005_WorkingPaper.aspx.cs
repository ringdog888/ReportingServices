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
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Bibliography;

namespace ReportingServices
{
    public partial class Rpt005_WorkingPaper : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            RptTool RptTool = new RptTool();
            //string EinB64 = Request.QueryString["EinB64"]?? "Q2xhc3M9MSZHcm91cD0wJkRlcHRJRD0wMDcmT3B0aW9uPTAmTDFJRD0=";
            string EinB64 = Request.QueryString["EinB64"] ?? "Z3VpZGlkPTAwNTI2QzEyLUI0QUQtNDRCOS04NjM0LUM3RkYwMzI5QTdCMCZEZXB0SUQ9QjA5NDA1MDAwMCZtQ2xhc3M9MA==";
            byte[] b = Convert.FromBase64String(EinB64);
            string bStr = System.Text.Encoding.UTF8.GetString(b);
            string[] baseStr = bStr.Split('&');
            string className = "";
            string mClass = "";
            string mGroup = "";
            string mDeptID = ""; 
            string mOption = "";
            string mL1ID = "";
            string guidid = "";
            string connectionString = RptTool.LoadCmdStr("\\\\Database\\\\Project\\\\BPM\\\\BPMPro\\\\Connection\\\\Audit.xdbc.xmf");
            //string connectionString = "Password=sa;Persist Security Info=True;User ID=sa;Initial Catalog=Audit;Data Source=WIN-7AVURFP25PL;";

            //Base64解碼
            foreach (string qstr in baseStr)
            {
                string[] qs = qstr.Split('=');
                if (qs[0] == "guidid") guidid = qs[1];
                if (qs[0] == "DeptID") mDeptID = qs[1];
                if (qs[0] == "mClass") mClass = qs[1];
            }
            /*
            mGroup = mGroup.Trim();
            mOption = mOption.Trim();
            mL1ID = mL1ID.Trim();*/
            mClass = mClass.Trim();
            mDeptID = mDeptID.Trim();
            guidid = guidid.Trim();
            string[] mStr = connectionString.Split(';');
            string mConStr = connectionString;

            /*foreach (string str in mStr)
            {
                string[] qs = str.Split('=');

                if (qs[0] == "Password") mConStr = mConStr + "Pwd=" + qs[1] + ";";
                if (qs[0] == "User ID") mConStr = mConStr + "Uid=" + qs[1] + ";";
                if (qs[0] == "Initial Catalog") mConStr = mConStr + "Database=" + qs[1] + ";";
                if (qs[0] == "Data Source") mConStr = mConStr + "Data Source=" + qs[1] + ";";

            }
            */
 
            DataTable table = new DataTable();

            Dictionary<string, string> dic = new Dictionary<string, string>();
            dic.Add("@guidid", guidid);
            dic.Add("@DeptID", mDeptID);
            table= RptTool.ExecSqlQueryParameters(connectionString,
                 @"select defaultname AS L1Name from audit_system_workingpaper_list
  left join audit_system_workingpaper_applicable_division 
  on audit_system_workingpaper_list.guidid=audit_system_workingpaper_applicable_division.working_paperid
  where groupid=@DeptID and enabled=1 and guidid=@guidid", dic);

            /*sqlcmd.Parameters.Add("@AuditGroup", SqlDbType.Int);
            sqlcmd.Parameters.Add("@AuditClass", SqlDbType.Int);
            sqlcmd.Parameters.Add("@ProjectOption", SqlDbType.Int);
            sqlcmd.Parameters.Add("@L1ID", SqlDbType.Int);*/


            if (mClass == "0")
            {
                className = "一般";
            }
            else
            {
                className = "專案";
            }

            if (table.Rows.Count == 0)
            {
                Dictionary<string, string> dic2 = new Dictionary<string, string>();
                table = RptTool.ExecSqlQueryParameters(connectionString, "SELECT 1,'' as L1Name",dic2);            
            }

            Rpt005WorkingPaper.GeneratedClass Rpt = new Rpt005WorkingPaper.GeneratedClass();
            string tempFile = Server.MapPath(".") + Guid.NewGuid() + ".docx";
            Rpt.className = className;
            Rpt.dt = table;
            Rpt.CreatePackage(tempFile);
            byte[] buff = File.ReadAllBytes(tempFile);
            File.Delete(tempFile);
            string fileName = "查核工作底稿封面.docx";
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