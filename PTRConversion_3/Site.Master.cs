using System;
using System.IO.Compression;
using System.Web.UI;

namespace PTRConversion_3
{
    public partial class SiteMaster : MasterPage
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            
        }

        protected void BtnDownloadArchives_Click(object sender, EventArgs e)
        {
            System.IO.File.Delete(Server.MapPath("~/Archive.zip"));
            string startPath = Server.MapPath("~/Archives/");
            string zipPath = Server.MapPath("~/Archive.zip");
            ZipFile.CreateFromDirectory(startPath, zipPath);
            Response.ContentType = "application/zip";
            Response.AppendHeader("Content-Disposition", "attachment; filename=Archive.zip");
            Response.TransmitFile(Server.MapPath("~/Archive.zip"));
            Response.End();
            System.IO.File.Delete(Server.MapPath("~/Archive.zip"));
        }
    }
}