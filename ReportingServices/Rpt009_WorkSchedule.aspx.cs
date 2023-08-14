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
using NPOI.Util;
using NPOI.XSSF.UserModel;
using DocumentFormat.OpenXml.Spreadsheet;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using CellType = NPOI.SS.UserModel.CellType;
using System.Linq;
using NPOI.SS.Formula.Functions;
using DocumentFormat.OpenXml.Office2016.Excel;
using NPOI.XWPF.UserModel;
using DocumentFormat.OpenXml.Drawing.Charts;
using System.Reflection;
using DataTable = System.Data.DataTable;

namespace ReportingServices
{
    public partial class Rpt009_WorkSchedule : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            RptTool RptTool = new RptTool();
            string EinB64 = Request.QueryString["EinB64"] ?? "";
            byte[] b = Convert.FromBase64String(EinB64);
            string bStr = System.Text.Encoding.UTF8.GetString(b);
            string[] baseStr = bStr.Split('&');
            string Year = "";
            string DeptID = "";
            string Month = "";
            string connectionString = RptTool.LoadCmdStr("\\\\Database\\\\Project\\\\BPM\\\\BPMPro\\\\Connection\\\\Audit.xdbc.xmf");
            //string connectionString = "Password=New@type;Persist Security Info=True;User ID=sa;Initial Catalog=Audit;Data Source=192.168.7.120;";

            foreach (string qstr in baseStr)
            {
                string[] qs = qstr.Split('=');
                if (qs[0] == "Year") Year = qs[1];
                if (qs[0] == "DeptID") DeptID = qs[1];
                if (qs[0] == "Month") Month = qs[1];
            }
            Rpt009.GeneratedClass Rpt = new Rpt009.GeneratedClass();
            Rpt.connectionString = connectionString;
            Rpt.baseStr = baseStr;
            byte[] excelData= Rpt.CreatePackage();
            Response.AddHeader("Content-Disposition", string.Format("attachment; filename=" + Year + "-" + Month + "出差日程表.xlsx"));
            // 輸出檔案
            Response.BinaryWrite(excelData);
            Response.End();
        }
    }
}