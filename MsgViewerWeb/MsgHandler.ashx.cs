using System;
using System.IO;
using System.Reflection;
using System.Web;
using System.Web.SessionState;
using DocumentServices.Modules.Readers.MsgReader;

namespace MsgViewerWeb
{
    /// <summary>
    /// Summary description for MsgHandler
    /// </summary>
    public class MsgHandler : IHttpHandler, IRequiresSessionState
    {
        private const string VirtualMessagesDir = @"~/FileSystem";

        public string RootMessagesDir { get; set; }

        public void ProcessRequest(HttpContext context)
        {
            RootMessagesDir = context.Server.MapPath(VirtualMessagesDir);
            var msgFileName = context.Request.QueryString.Get("file");
            var msgFileFullName = Path.Combine(RootMessagesDir, msgFileName);
            var attachmentFileName = context.Request.QueryString.Get("attachment");

            if (!string.IsNullOrEmpty(attachmentFileName))
            {
                ServerAttachment(msgFileName, attachmentFileName, context);
                return;
            }

            ServeMsgFile(msgFileFullName, context, msgFileName);
        }

        private void ServerAttachment(string msgFileName, string attachedfileName, HttpContext context)
        {
            var attachmentFolder = Path.Combine(RootMessagesDir, "messages", context.Session.SessionID, Path.GetFileNameWithoutExtension(msgFileName));
            var attachedfileFullName = Path.Combine(attachmentFolder, attachedfileName);
            var ext = Path.GetExtension(attachedfileFullName);

            if (ext.ToLower() == ".msg")
            {
                ServeMsgFile(attachedfileFullName, context, attachedfileName);
                return;
            }
            if (File.Exists(attachedfileFullName))
            {
                context.Response.ContentType = MimeExtensionHelper.GetMimeType(attachedfileFullName);
                context.Response.TransmitFile(attachedfileFullName);
                return;
            }

            throw new HttpException(404, "File does not exist.");
        }

        private void ServeMsgFile(string fileFullName, HttpContext context, string fileName)
        {
            try
            {
                if (File.Exists(fileFullName))
                {
                    var msgFolder = Path.Combine(RootMessagesDir, "messages", context.Session.SessionID, Path.GetFileNameWithoutExtension(fileName));

                    if (!Directory.Exists(msgFolder))
                    {
                        Directory.CreateDirectory(msgFolder);
                    }

                    var emailReader = MessageReaderFactory.CreateMessageReader(context);

                    var files = emailReader.ExtractToFolder(fileFullName, msgFolder);

                    if (files.Length > 0)
                    {
                        //always extracts it to email.html
                        var email = files[0];

                        context.Response.ContentType = MimeExtensionHelper.GetMimeType(email);
                        context.Response.TransmitFile(email);
                        return;
                    }
                    throw new HttpException(404, "File could not be converted.");
                }
                throw new HttpException(404, "File does not exist.");
            }
            catch (HttpException e)
            {
                context.Response.StatusCode = e.GetHttpCode();
                context.Response.Status = e.Message;
                context.Response.End();
            }
        }

        public bool IsReusable
        {
            get
            {
                return false;
            }
        }
    }

    public static class MimeExtensionHelper
    {
        static object locker = new object();
        static object mimeMapping;
        static MethodInfo getMimeMappingMethodInfo;

        static MimeExtensionHelper()
        {
            Type mimeMappingType = Assembly.GetAssembly(typeof(HttpRuntime)).GetType("System.Web.MimeMapping");
            if (mimeMappingType == null)
                throw new SystemException("Couldn't find MimeMapping type");
            getMimeMappingMethodInfo = mimeMappingType.GetMethod("GetMimeMapping", BindingFlags.Static | BindingFlags.NonPublic | BindingFlags.Public);
            if (getMimeMappingMethodInfo == null)
                throw new SystemException("Couldn't find GetMimeMapping method");
            if (getMimeMappingMethodInfo.ReturnType != typeof(string))
                throw new SystemException("GetMimeMapping method has invalid return type");
            if (getMimeMappingMethodInfo.GetParameters().Length != 1 && getMimeMappingMethodInfo.GetParameters()[0].ParameterType != typeof(string))
                throw new SystemException("GetMimeMapping method has invalid parameters");
        }
        public static string GetMimeType(string filename)
        {
            lock (locker)
                return (string)getMimeMappingMethodInfo.Invoke(mimeMapping, new object[] { filename });
        }
    }

}

