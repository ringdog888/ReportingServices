using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

public partial class ZipDownload : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        int i = 0;
        string ZipPath = Server.MapPath(".") + Guid.NewGuid() + ".zip";
        List<string> FileList = new List<string>();
        string FileName = Request["FileName"];
        string DomainURL = Request.Url.AbsoluteUri;

        DomainURL = DomainURL.Substring(0, DomainURL.LastIndexOf('/') + 1);
        
        while (Request["FileList" + i.ToString()] != null)
        {
            FileList.Add(DomainURL + Request["FileList" + i.ToString()]);
            i++;
        }

        ZipTool zip = new ZipTool();
        zip.CreateZipFile(FileList, ZipPath);
        byte[] buff = File.ReadAllBytes(ZipPath);
        File.Delete(ZipPath);

        if (Request.Browser.Browser == "IE")
        {
            FileName = Server.UrlPathEncode(FileName);
        }
        Response.Clear();
        Response.ContentType = "application/octet-stream";
        Response.AddHeader("content-disposition", "attachment;filename=" + FileName + ".zip");
        Response.BinaryWrite(buff);
        Response.End();
    }
}