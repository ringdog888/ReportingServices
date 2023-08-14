using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;
using Newtonsoft.Json;
using NPOI.XSSF.UserModel;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using System.Data;
using System;
using DocumentFormat.OpenXml.Packaging;
using NPOI.SS.Util;
using System.Collections.Generic;
using System.IO;

namespace Rpt009
{
    public class GeneratedClass
    {
        //Creates Report Tool
       ReportingServices.RptTool RptTool = new ReportingServices.RptTool();

        // Data Source
        public string connectionString { get; set; }
        public string[] baseStr { get; set; }

        public byte[] CreatePackage()
        { 

             string Year = "";
            string DeptID = "";
            string Month = "";

            //Base64解碼
            foreach (string qstr in baseStr)
            {
                string[] qs = qstr.Split('=');
                if (qs[0] == "Year") Year = qs[1];
                if (qs[0] == "DeptID") DeptID = qs[1];
                if (qs[0] == "Month") Month = qs[1];
            }
            string[] DeptIDArr = DeptID.Split(',');
            string TM = Year + "-" + Month + "-01";
            XSSFWorkbook workbook = new XSSFWorkbook();
            foreach (string dept in DeptIDArr)
            {
                Dictionary<string, string> dic = new Dictionary<string, string>();
                dic.Add("@TM", TM);
                dic.Add("@DeptID", dept);
                DataTable dt = RptTool.ExecSqlQueryParameters(connectionString,
                        @"select audit_system_auditplan_list.guidid,CASE WHEN dbo.Auditfn_GetLang_plantype_list(audit_system_auditplan_list.plantypeid,'zh-tw')='一般查核'
                        THEN '一般查核'
                        ELSE '專案查核' END
                        AS plantype,accountid,
                        REPLACE(ISNULL(dbo.Orgfn_GetLangMemName(accountid,'zh-tw'),N'No Name'),' ','') AS MemberName,
                        (select DeptID from [dbo].[F7Organ_View_CurrMember] where AccountID=audit_system_auditplan_members.accountid ) AS DeptID,
                        (select DeptName from [dbo].[F7Organ_View_CurrMember] where AccountID=audit_system_auditplan_members.accountid ) AS DeptName,
                        startdate,enddate,audit_system_auditplan_depts.deptid AS Audit_DeptID,REPLACE(dbo.Orgfn_GetLangDeptName(audit_system_auditplan_depts.deptid,'zh-tw'),'分行','') AS Audit_DeptName
                        from audit_system_auditplan_members
                        left join audit_system_auditplan_list
                        on audit_system_auditplan_list.guidid=audit_system_auditplan_members.planid
                        left join audit_system_auditplan_depts
                        on audit_system_auditplan_depts.planid=audit_system_auditplan_members.planid
                        where startdate is not null and enddate is not null and dbo.Auditfn_GetLang_plantype_list(audit_system_auditplan_list.plantypeid,'zh-tw') is not null
                        and (startdate between @TM and CAST(dateadd(day ,-1, dateadd(m, datediff(m,0,@TM)+1,0)) AS Date) or enddate between @TM and CAST(dateadd(day ,-1, dateadd(m, datediff(m,0,@TM)+1,0)) AS Date))
                        and (select COUNT(1) from [dbo].[F7Organ_View_CurrMember] where AccountID=audit_system_auditplan_members.accountid and DeptID 
                        in('H00507D200','H00507D300','H00507D100'))>0 and (select COUNT(1) from [dbo].[F7Organ_View_CurrMember] where AccountID=audit_system_auditplan_members.accountid and DeptID 
                        =@DeptID)>0
                        order by startdate,audit_system_auditplan_list.guidid ", dic);

                DataTable dt_Count = RptTool.ExecSqlQueryParameters(connectionString,
                        @"select COUNT(1) AS CNT from (
                        select accountid
                        from audit_system_auditplan_members
                        left join audit_system_auditplan_list
                        on audit_system_auditplan_list.guidid=audit_system_auditplan_members.planid
                        left join audit_system_auditplan_depts
                        on audit_system_auditplan_depts.planid=audit_system_auditplan_members.planid
                        where startdate is not null and enddate is not null and dbo.Auditfn_GetLang_plantype_list(audit_system_auditplan_list.plantypeid,'zh-tw') is not null
                        and (startdate between @TM and CAST(dateadd(day ,-1, dateadd(m, datediff(m,0,@TM)+1,0)) AS Date) or enddate between @TM and CAST(dateadd(day ,-1, dateadd(m, datediff(m,0,@TM)+1,0)) AS Date))
                        and (select COUNT(1) from [dbo].[F7Organ_View_CurrMember] where AccountID=audit_system_auditplan_members.accountid and DeptID 
                        in('H00507D200','H00507D300','H00507D100'))>0 and (select COUNT(1) from [dbo].[F7Organ_View_CurrMember] where AccountID=audit_system_auditplan_members.accountid and DeptID 
                        =@DeptID)>0 group by accountid) AS T ", dic);
                string sheetname = "";
                if (dept == "H00507D200")
                {
                    sheetname = "稽查二部";
                }
                else if (dept == "H00507D300")
                {
                    sheetname = "規劃部";
                }
                else if (dept == "H00507D100")
                {
                    sheetname = "稽查一部";
                }
                else
                {
                    sheetname = "A";
                }
                XSSFSheet sheet = (XSSFSheet)workbook.CreateSheet(sheetname);
                XSSFFont ContentFont = (XSSFFont)workbook.CreateFont();
                ContentFont.FontHeightInPoints = (short)12;
                ContentFont.FontName = "Arial";

                XSSFFont Titlefont = (XSSFFont)workbook.CreateFont();
                Titlefont.FontHeightInPoints = (short)12;
                Titlefont.FontName = "Arial";
                Titlefont.Boldweight = (short)FontBoldWeight.Bold;

                XSSFCellStyle ContentStyle = (XSSFCellStyle)workbook.CreateCellStyle();
                ContentStyle.SetFont(ContentFont);
                ContentStyle.DataFormat = workbook.CreateDataFormat().GetFormat("text");
                ContentStyle.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
                ContentStyle.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
                ContentStyle.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
                ContentStyle.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
                ContentStyle.Alignment = HorizontalAlignment.Center;
                ContentStyle.VerticalAlignment = VerticalAlignment.Center;

                XSSFCellStyle NormalStyle = (XSSFCellStyle)workbook.CreateCellStyle();
                NormalStyle.SetFont(Titlefont);
                var color_Normal = new XSSFColor(new byte[] { 102, 204, 102 });
                NormalStyle.SetFillForegroundColor(color_Normal);
                NormalStyle.FillPattern = FillPattern.SolidForeground;
                NormalStyle.DataFormat = workbook.CreateDataFormat().GetFormat("text");
                NormalStyle.WrapText = true;
                NormalStyle.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
                NormalStyle.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
                NormalStyle.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
                NormalStyle.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
                NormalStyle.Alignment = HorizontalAlignment.Center;
                NormalStyle.VerticalAlignment = VerticalAlignment.Center;

                XSSFCellStyle ProjectStyle = (XSSFCellStyle)workbook.CreateCellStyle();
                ProjectStyle.SetFont(Titlefont);
                var color_Project = new XSSFColor(new byte[] { 255, 255, 0 });
                ProjectStyle.SetFillForegroundColor(color_Project);
                ProjectStyle.FillPattern = FillPattern.SolidForeground;
                ProjectStyle.DataFormat = workbook.CreateDataFormat().GetFormat("text");
                ProjectStyle.WrapText = true;
                ProjectStyle.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
                ProjectStyle.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
                ProjectStyle.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
                ProjectStyle.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
                ProjectStyle.Alignment = HorizontalAlignment.Center;
                ProjectStyle.VerticalAlignment = VerticalAlignment.Center;

                int year = Int32.Parse(Year);
                int month = Int32.Parse(Month);

                DateTime startDate = new DateTime(year, month, 1);
                DateTime endDate = startDate.AddMonths(1).AddDays(-1);
                // Iterate through each day of the month
                CellRangeAddress region = new CellRangeAddress(0, 2, 0, 2);
                sheet.AddMergedRegion(region);
                sheet.CreateRow(0);
                sheet.GetRow(0).CreateCell(0, CellType.String).SetCellValue("姓名╲日期");
                sheet.GetRow(0).GetCell(0).CellStyle = ContentStyle;
                CellRangeAddress region2 = new CellRangeAddress(0, 0, 3, DateTime.DaysInMonth(year, month));
                sheet.AddMergedRegion(region2);
                sheet.GetRow(0).CreateCell(3, CellType.String).SetCellValue(dt_Count.Rows[0]["CNT"].ToString() + "員");
                sheet.GetRow(0).GetCell(3).CellStyle = ContentStyle;
                int i = 3;
                sheet.CreateRow(1);
                sheet.CreateRow(2);
                for (DateTime date = startDate; date <= endDate; date = date.AddDays(1))
                {
                    sheet.GetRow(1).CreateCell(i, CellType.String).SetCellValue(i - 2);
                    sheet.GetRow(1).GetCell(i).CellStyle = ContentStyle;
                    sheet.GetRow(2).CreateCell(i, CellType.String).SetCellValue(GetChineseDayOfWeek(date.DayOfWeek));
                    sheet.GetRow(2).GetCell(i).CellStyle = ContentStyle;
                    i++;
                }
                i = 3;
                List<string> accountArr = new List<string>();
                foreach (DataRow rows in dt.Rows)
                {
                    if (!accountArr.Contains(rows["accountid"].ToString()))
                    {

                        CellRangeAddress regiontemp = new CellRangeAddress(i, i, 0, 2);
                        sheet.AddMergedRegion(regiontemp);
                        sheet.CreateRow(i);
                        accountArr.Add(rows["accountid"].ToString());
                        sheet.GetRow(i).CreateCell(0, CellType.String).SetCellValue(rows["MemberName"].ToString());
                        sheet.GetRow(i).GetCell(0).CellStyle = ContentStyle;
                        RegionStyleSetting(sheet.GetRow(i), regiontemp, ContentStyle);
                        i++;
                    }
                    DateTime startdate = DateTime.Parse(rows["startdate"].ToString());
                    DateTime enddate = DateTime.Parse(rows["enddate"].ToString());
                    int dayS = startdate.Day;
                    int dayE = enddate.Day;
                    int monthS = enddate.Month;
                    int monthE = enddate.Month;
                    int BarStart = dayS + 2;
                    int BarEnd = dayE + 2;
                    if (monthS < month)
                    {
                        BarStart = 3;
                    }
                    if (monthE > month)
                    {
                        BarEnd = DateTime.DaysInMonth(year, month) + 3;
                    }

                    int index = accountArr.IndexOf(rows["accountid"].ToString());
                    index = index + 3;
                    CellRangeAddress regiontempBar = new CellRangeAddress(index, index, BarStart, BarEnd);
                    sheet.AddMergedRegion(regiontempBar);
                    sheet.GetRow(index).CreateCell(BarStart, CellType.String).SetCellValue(rows["Audit_DeptName"].ToString());
                    if (rows["plantype"].ToString() == "一般查核")
                    {
                        sheet.GetRow(index).GetCell(BarStart).CellStyle = NormalStyle;
                        RegionStyleSetting(sheet.GetRow(index), regiontempBar, NormalStyle);
                    }
                    else
                    {
                        sheet.GetRow(index).GetCell(BarStart).CellStyle = ProjectStyle;
                        RegionStyleSetting(sheet.GetRow(index), regiontempBar, ProjectStyle);
                    }

                }
                for (int row = 3; row < i; row++)
                {
                    for (int col = 3; col < (3 + DateTime.DaysInMonth(year, month)); col++)
                    {
                        if (sheet.GetRow(row).GetCell(col) == null)
                        {
                            sheet.GetRow(row).CreateCell(col);
                            sheet.GetRow(row).GetCell(col).CellStyle = ContentStyle;
                        }
                    }
                }
                sheet.CreateRow(i + 2);
                sheet.GetRow(i + 2).CreateCell(2, CellType.String).SetCellValue("");
                sheet.GetRow(i + 2).GetCell(2).CellStyle = NormalStyle;
                sheet.GetRow(i + 2).CreateCell(3, CellType.String).SetCellValue("一般查核");
                sheet.GetRow(i + 2).GetCell(3).CellStyle = ContentStyle;
                sheet.GetRow(i + 2).CreateCell(4, CellType.String).SetCellValue("");
                sheet.GetRow(i + 2).GetCell(4).CellStyle = ProjectStyle;
                sheet.GetRow(i + 2).CreateCell(5, CellType.String).SetCellValue("專案查核");
                sheet.GetRow(i + 2).GetCell(5).CellStyle = ContentStyle;
            }
            byte[] excelData;
            using (MemoryStream ms = new MemoryStream())
            {

                workbook.Write(ms);
                excelData = ms.ToArray();
                workbook.Close();
                workbook = null;
            }
            return excelData;
        }

        static string GetChineseDayOfWeek(DayOfWeek dayOfWeek)
        {
            switch (dayOfWeek)
            {
                case DayOfWeek.Sunday:
                    return "日";
                case DayOfWeek.Monday:
                    return "一";
                case DayOfWeek.Tuesday:
                    return "二";
                case DayOfWeek.Wednesday:
                    return "三";
                case DayOfWeek.Thursday:
                    return "四";
                case DayOfWeek.Friday:
                    return "五";
                case DayOfWeek.Saturday:
                    return "六";
                default:
                    return string.Empty;
            }
        }
        static void RegionStyleSetting(IRow row, CellRangeAddress region, XSSFCellStyle cell_Style)
        {
            for (int colIndex = region.FirstColumn + 1; colIndex <= region.LastColumn; colIndex++)
            {
                row.CreateCell(colIndex, CellType.String);
                row.GetCell(colIndex).CellStyle = cell_Style;
            }
        }

        }
    }
