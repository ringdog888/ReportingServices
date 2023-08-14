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
    public partial class Rpt003_PracticeCard_3 : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            RptTool RptTool = new RptTool();
            string EinB64 = Request.QueryString["EinB64"];
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

            string sqlstr = "SELECT UniqueID, FullName, Birthday, CertfcateNo, AuditItem, Gender, sDay, eDay, bookDate, " +
                            "       YEAR(Birthday) - 1911 AS BirYear, MONTH(Birthday) AS BirMonth, DAY(Birthday) AS BirDay,  " +
                            "       YEAR(sDay) - 1911 AS SYear, MONTH(sDay) AS SMonth, DAY(sDay) AS S_Day,  " +
                            "       YEAR(eDay) - 1911 AS EYear, MONTH(eDay) AS EMonth, DAY(eDay) AS E_Day,  "+
                            "       YEAR(bookDate) AS BYear, MONTH(bookDate) AS BMonth, DAY(bookDate) AS BDay "+
                            "FROM   (SELECT TPM.UniqueID, TPM.FullName, TPM.Birthday, TPM.CertfcateNo, TPM.AuditItem, CASE Gender WHEN 0 THEN '女' WHEN 1 THEN '男' END AS Gender, " +
                            "               (SELECT TOP (1) sDate "+
                            "                  FROM eAudit_Training_MemStage "+
                            "                 WHERE (TPMID = TPM.UniqueID) AND (Exist = 1) "+
                            "                ORDER BY sDate) AS sDay, "+
                            "               (SELECT TOP (1) eDate "+
                            "                  FROM eAudit_Training_MemStage AS eAudit_Training_MemStage_1 "+
                            "                 WHERE (TPMID = TPM.UniqueID) AND (Exist = 1) "+
                            "                ORDER BY eDate DESC) AS eDay, TP.bookDate "+
                            "        FROM   eAudit_Training_PlanMem AS TPM INNER JOIN "+
                            "               eAudit_Training_Plans AS TP ON TPM.TPID = TP.UniqueID "+
                            "        WHERE      (TPM.TPID = " + ID + ")) AS t";

            DataTable dt = RptTool.ExecOleQuery(connectionString, sqlstr);

            Rpt003_PracticeCard3.GeneratedClass Rpt = new Rpt003_PracticeCard3.GeneratedClass();
            string tempFile = Server.MapPath(".") + Guid.NewGuid() + ".docx";
            Rpt.dt = dt;
            Rpt.CreatePackage(tempFile);
            byte[] buff = File.ReadAllBytes(tempFile);
            File.Delete(tempFile);
            string fileName = "實習證明書3.docx";
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