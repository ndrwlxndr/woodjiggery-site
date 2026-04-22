<%@ WebHandler Language="C#" Class="UploadHandler" %>

using System;
using System.IO;
using System.Web;

/// <summary>
/// Receives a multipart file upload and saves it to a folder under /images/.
/// Query string parameters:
///   dest  - required - subfolder path relative to /images/, e.g. "products/42"
///           Only alphanumeric characters, hyphens, underscores and forward slashes allowed.
/// Returns plain text: "OK|filename" on success, "ERROR|message" on failure.
/// </summary>
public class UploadHandler : IHttpHandler
{
    private static readonly string[] AllowedExtensions = { ".jpg", ".jpeg", ".png", ".gif", ".webp" };

    public void ProcessRequest(HttpContext ctx)
    {
        ctx.Response.ContentType = "text/plain";
        ctx.Response.TrySkipIisCustomErrors = true;

        try
        {
            // Validate dest parameter
            string dest = (ctx.Request.QueryString["dest"] ?? "").Trim().Replace('\\', '/');
            if (!IsValidRelativePath(dest))
            {
                ctx.Response.StatusCode = 400;
                ctx.Response.Write("ERROR|Invalid dest parameter");
                return;
            }

            if (ctx.Request.Files.Count == 0)
            {
                ctx.Response.StatusCode = 400;
                ctx.Response.Write("ERROR|No file received");
                return;
            }

            HttpPostedFile file = ctx.Request.Files[0];
            if (file == null || file.ContentLength == 0)
            {
                ctx.Response.StatusCode = 400;
                ctx.Response.Write("ERROR|Empty file");
                return;
            }

            // Validate extension
            string ext = Path.GetExtension(file.FileName).ToLowerInvariant();
            bool extOk = false;
            foreach (string a in AllowedExtensions) { if (ext == a) { extOk = true; break; } }
            if (!extOk)
            {
                ctx.Response.StatusCode = 400;
                ctx.Response.Write("ERROR|File type not allowed: " + HtmlEncode(ext));
                return;
            }

            // Build destination path
            string safeName = Path.GetFileName(file.FileName); // strips any directory component
            string imagesRoot = ctx.Server.MapPath("/images");
            string destFolder = dest.Length > 0
                ? Path.Combine(imagesRoot, dest.Replace('/', Path.DirectorySeparatorChar))
                : imagesRoot;

            // Ensure destination folder exists (create up to two levels deep)
            if (!Directory.Exists(destFolder))
                Directory.CreateDirectory(destFolder);

            string destPath = Path.Combine(destFolder, safeName);
            file.SaveAs(destPath);

            ctx.Response.Write("OK|" + safeName);
        }
        catch (Exception ex)
        {
            ctx.Response.TrySkipIisCustomErrors = true;
            ctx.Response.StatusCode = 500;
            ctx.Response.Write("ERROR|" + ex.Message);
        }
    }

    private static bool IsValidRelativePath(string path)
    {
        if (path.Length == 0) return true; // root /images/ is fine
        if (path.Contains("..")) return false;
        if (path.Contains(":")) return false;
        foreach (char c in path)
        {
            if (!char.IsLetterOrDigit(c) && c != '/' && c != '-' && c != '_')
                return false;
        }
        return true;
    }

    private static string HtmlEncode(string s)
    {
        return HttpUtility.HtmlEncode(s);
    }

    public bool IsReusable { get { return false; } }
}
