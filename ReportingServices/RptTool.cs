using System;
using System.Collections.Generic;
using System.IO;
using System.Web;
using System.Data;
using System.Data.SqlClient;
using System.Data.OleDb;
using Microsoft.Win32;
using NAWXDBCINFOIOLib;
using System.Web.UI;
using System.Web.UI.WebControls;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Spreadsheet;
using Run = DocumentFormat.OpenXml.Wordprocessing.Run;
using Break = DocumentFormat.OpenXml.Wordprocessing.Break;

namespace ReportingServices
{
    public class RptTool
    {
        public Paragraph GetBreakValues()
        {
            Paragraph paragraph = new Paragraph() { RsidParagraphMarkRevision = "00EE1A66", RsidParagraphAddition = "00704F19", RsidParagraphProperties = "00292C3A", RsidRunAdditionDefault = "00292C3A" };

            ParagraphProperties paragraphProperties = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties = new ParagraphMarkRunProperties();
            RunFonts runFonts120 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };

            paragraphMarkRunProperties.Append(runFonts120);

            paragraphProperties.Append(paragraphMarkRunProperties);

            Run run = new Run();
            Break break1 = new Break() { Type = BreakValues.Page };

            run.Append(break1);

            paragraph.Append(paragraphProperties);
            paragraph.Append(run);
            return paragraph;
        }

        public DataTable ExecOleQuery(string connString, string CommandText)
        {
            DataSet ds = new DataSet();
            DataTable dt = new DataTable();
            using (OleDbConnection Conn = new OleDbConnection(connString))
            {
                Conn.Open();
                OleDbDataAdapter adapter = new OleDbDataAdapter();

                adapter.SelectCommand = new OleDbCommand(CommandText, Conn);
                adapter.Fill(ds, "ds");
                Conn.Close();
            }
            if (ds.Tables.Count > 0) {
                dt = ds.Tables[0];
            }
            return dt;
        }
        public DataTable ExecSqlQueryParameters(string connString, string CommandText, Dictionary<string, string> Map)
        {
            DataTable dt = new DataTable();
            using (SqlConnection Conn = new SqlConnection(connString))
            {
                Conn.Open();
                SqlCommand cmd = new SqlCommand();
                cmd.Connection = Conn;
                cmd.CommandText = CommandText;
                foreach (KeyValuePair<string, string> kvp in Map)
                {
                    SqlParameter parameter = new SqlParameter(kvp.Key, SqlDbType.NVarChar, 65535);
                    if (kvp.Value == "")
                    {
                        parameter.Value = DBNull.Value;
                    }
                    else
                    {
                        parameter.Value = kvp.Value;
                    }
                    cmd.Parameters.Add(parameter);
                }
                SqlDataAdapter dAdpter = new SqlDataAdapter(cmd);
                dAdpter.Fill(dt);
                Conn.Close();
            }
            return dt;
        }

        public int ExecNonQuery(string connString, string CommandText)
        {
            int rVal;
            using (OleDbConnection Conn = new OleDbConnection(connString))
            {
                Conn.Open();
                OleDbCommand cmd = new OleDbCommand(CommandText, Conn);
                rVal = cmd.ExecuteNonQuery();
                Conn.Close();
            }
            return rVal;
        }

        public int ExecSqlNonQuery(string connString, string CommandText)
        {
            int rVal;
            using (SqlConnection Conn = new SqlConnection(connString))
            {
                Conn.Open();
                SqlCommand cmd = new SqlCommand(CommandText, Conn);
                rVal = cmd.ExecuteNonQuery();
                Conn.Close();
            }
            return rVal;
        }

        /// <summary>
        /// 取得AutoWEB資料庫連線字串
        /// </summary>
        /// <param name="xdbcPath">xdbc.Xmf所在路徑</param>
        /// <returns></returns>
        public string LoadCmdStr(String xdbcPath)
        {
            string FileName = "";
            const string userRoot = "HKEY_LOCAL_MACHINE";
            const string subkey = "Software\\NewType\\AutoWeb.Net";
            const string keyName = userRoot + "\\" + subkey;
            string Path = (string)Registry.GetValue(keyName, "Root", -1);
            Path = Path.Replace("\\", "\\\\");
            XdbcInfoIO objXdbc = new XdbcInfoIO();
            FileName = Path + xdbcPath;
            objXdbc.LoadFile(FileName, "");
            string connectionString = objXdbc.XdbcConnection.sMsSqlConnectString;
            return connectionString;
        }
    }
}