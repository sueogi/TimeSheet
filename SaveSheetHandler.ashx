<%@ WebHandler Language="C#" Class="SaveSheetHandler" %>

using System;
using System.IO;
using System.Web;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

public class SaveSheetHandler : IHttpHandler, System.Web.SessionState.IRequiresSessionState
{

    public void ProcessRequest(HttpContext context)
    {        
        StreamReader stream = new StreamReader(context.Request.InputStream);
        string text = stream.ReadToEnd();        
        SheetDetail sheet = JsonConvert.DeserializeObject(text, typeof(SheetDetail)) as SheetDetail;
        sheet.Account = context.Session["account"].ToString();
        string result = string.Empty;
        result = SaveTimeSheetData.SaveData(sheet);
        
        context.Response.ContentType = "text/plain";       
        context.Response.Write(result);
    }

    public bool IsReusable
    {
        get
        {
            return false;
        }
    }
}