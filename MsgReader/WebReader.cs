using System;
using System.Collections.Generic;
using System.IO;
using System.Web;
using DocumentServices.Modules.Readers.MsgReader.Outlook;

namespace DocumentServices.Modules.Readers.MsgReader
{
    public class WebReader : IReader
    {

        public WebReader(string msgHandlerUrl, string applicationPath, string virtualPath)
        {
            MsgHandlerUrl = msgHandlerUrl;
            ApplicationPath = applicationPath;
            VirtualPath = virtualPath;
        }

        public string MsgHandlerUrl { get; private set; }
        protected string VirtualPath { get; private set; }
        public string ApplicationPath { get; private set; }

        public string[] ExtractToFolder(string inputFile, string outputFolder, bool hyperlinks = false)
        {
            outputFolder = FileManager.CheckForBackSlash(outputFolder);
            _errorMessage = string.Empty;

            try
            {
                using (var messageStream = File.Open(inputFile, FileMode.Open, FileAccess.Read))
                {
                    // Read MSG file from a stream
                    using (var message = new Storage.Message(messageStream))
                    {
                        var result = new List<string>();
                        // Determine the name for the E-mail body
                        var eMailFileName = Path.Combine(outputFolder,
                                                         "email" + (message.BodyHtml != null ? ".html" : ".txt"));
                        result.Add(eMailFileName);

                        // We first always check if there is a HTML body
                        var body = message.BodyHtml;
                        var htmlBody = true;
                        if (body == null)
                        {
                            // When not found try to get the text body
                            body = message.BodyText;
                            htmlBody = false;
                        }

                        var attachments = string.Empty;

                        foreach (var attachment in message.Attachments)
                        {
                            var fileName = string.Empty;
                            if (attachment.GetType() == typeof(Storage.Attachment))
                            {
                                var attach = (Storage.Attachment)attachment;
                                fileName =
                                    FileManager.FileExistsMakeNew(outputFolder +
                                                                  FileManager.RemoveInvalidFileNameChars(attach.FileName)
                                                                             .Replace("&", "[and]"));
                                File.WriteAllBytes(fileName, attach.Data);

                                // When we find an in-line attachment we have to replace the CID tag inside the HTML body
                                // with the name of the in-line attachment. But before we do this we check if the CID exists.
                                // When the CID does not exists we treat the in-line attachment as a normal attachment
                                if (htmlBody && !string.IsNullOrEmpty(attach.ContentId) && body.Contains(attach.ContentId))
                                {
                                    body = body.Replace("cid:" + attach.ContentId, GetRelativePathFromAbsolutePath(fileName));
                                    continue;
                                }

                                result.Add(fileName);

                            }
                            else if (attachment.GetType() == typeof(Storage.Message))
                            {
                                var msg = (Storage.Message)attachment;
                                fileName =
                                    FileManager.FileExistsMakeNew(outputFolder +
                                                                  FileManager.RemoveInvalidFileNameChars(msg.Subject) +
                                                                  ".msg").Replace("&", "[and]");
                                result.Add(fileName);
                                msg.Save(fileName);
                            }

                            //var ext = Path.GetExtension(fileName);

                            if (attachments == string.Empty)
                                //attachments = BuildAnchor(fileName, !string.IsNullOrEmpty(ext) && ext.ToLower() == ".msg" ? MsgHandlerUrl : string.Empty); // Path.GetFileName(fileName);
                                attachments = BuildAnchor(fileName, MsgHandlerUrl);
                            else
                                //attachments += ", " + BuildAnchor(fileName, !string.IsNullOrEmpty(ext) && ext.ToLower() == ".msg" ? MsgHandlerUrl : string.Empty); // Path.GetFileName(fileName);
                                attachments += ", " + BuildAnchor(fileName, MsgHandlerUrl);
                        }

                        string outlookHeader;

                        if (htmlBody)
                        {
                            // Add an outlook style header into the HTML body.
                            // Change this code to the language you need. 
                            // Currently it is written in ENGLISH
                            outlookHeader =
                                "<TABLE cellSpacing=0 cellPadding=0 width=\"100%\" border=0 style=\"font-family: 'Times New Roman'; font-size: 12pt;\"\\>" + Environment.NewLine +
                                "<TR><TD valign=\"top\" style=\"height: 18px; width: 100px \"><STRONG>From:</STRONG></TD><TD valign=\"top\" style=\"height: 18px\">" + GetEmailSender(message) + "</TD></TR>" + Environment.NewLine +
                                "<TR><TD valign=\"top\" style=\"height: 18px; width: 100px \"><STRONG>To:</STRONG></TD><TD valign=\"top\" style=\"height: 18px\">" + GetEmailRecipients(message, Storage.Recipient.RecipientType.To) + "</TD></TR>" + Environment.NewLine +
                                "<TR><TD valign=\"top\" style=\"height: 18px; width: 100px \"><STRONG>Sent on:</STRONG></TD><TD valign=\"top\" style=\"height: 18px\">" + (message.SentOn != null ? ((DateTime)message.SentOn).ToString("dd-MM-yyyy HH:mm:ss") : string.Empty) + "</TD></TR>";

                            // CC
                            var cc = GetEmailRecipients(message, Storage.Recipient.RecipientType.Cc);
                            if (cc != string.Empty)
                                outlookHeader += "<TR><TD valign=\"top\" style=\"height: 18px; width: 100px \"><STRONG>CC:</STRONG></TD><TD style=\"height: 18px\">" + cc + "</TD></TR>" + Environment.NewLine;

                            // Subject
                            outlookHeader += "<TR><TD valign=\"top\" style=\"height: 18px; width: 100px \"><STRONG>Subject:</STRONG></TD><TD style=\"height: 18px\">" + message.Subject + "</TD></TR>" + Environment.NewLine;

                            // Empty line
                            outlookHeader += "<TR><TD colspan=\"2\" style=\"height: 18px\">&nbsp</TD></TR>" + Environment.NewLine;

                            // Attachments
                            if (attachments != string.Empty)
                                outlookHeader += "<TR><TD valign=\"top\" style=\"height: 18px; width: 100px \"><STRONG>Attachments:</STRONG></TD><TD style=\"height: 18px\">" + attachments + "</TD></TR>" + Environment.NewLine;

                            //  End of table + empty line
                            outlookHeader += "</TABLE><BR>" + Environment.NewLine;

                            body = InjectOutlookHeader(body, outlookHeader);
                        }
                        else
                        {
                            // Add an outlook style header into the Text body. 
                            // Change this code to the language you need. 
                            // Currently it is written in ENGLISH
                            outlookHeader =
                                "From:\t\t" + GetEmailSender(message) + Environment.NewLine +
                                "To:\t\t" + GetEmailRecipients(message, Storage.Recipient.RecipientType.To) + Environment.NewLine +
                                "Sent on:\t" + (message.SentOn != null ? ((DateTime)message.SentOn).ToString("dd-MM-yyyy HH:mm:ss") : string.Empty) + Environment.NewLine;

                            // CC
                            var cc = GetEmailRecipients(message, Storage.Recipient.RecipientType.Cc);
                            if (cc != string.Empty)
                                outlookHeader += "CC:\t\t" + cc + Environment.NewLine;

                            outlookHeader += "Subject:\t" + message.Subject + Environment.NewLine + Environment.NewLine;

                            // Attachments
                            if (attachments != string.Empty)
                                outlookHeader += "Attachments:\t" + attachments + Environment.NewLine + Environment.NewLine;

                            body = outlookHeader + body;
                        }

                        // Write the body to a file
                        File.WriteAllText(eMailFileName, body);
                        return result.ToArray();
                    }
                }
            }
            catch (Exception e)
            {
                //if (message != null)
                //    message.Dispose();
                _errorMessage = GetInnerException(e);
                return new string[0];
            }
        }

        #region GetErrorMessage
        /// <summary>
        /// Get the last know error message. When the string is empty there are no errors
        /// </summary>
        /// <returns></returns>
        public string GetErrorMessage()
        {
            return _errorMessage;
        }
        #endregion

        private string BuildAnchor(string fileFullName, string customHandler = "")
        {
            var fileName = Path.GetFileName(fileFullName);
            if (string.IsNullOrEmpty(customHandler))
            {
                return string.Format(@"<a href=""{0}"">{1}</a> ", GetRelativePathFromAbsolutePath(fileFullName, true), fileName);
            }
            return string.Format(@"<a href=""{0}"">{1}</a> ", customHandler + "&attachment=" + fileName, fileName);
        }

        private string GetRelativePathFromAbsolutePath(string path, bool encode = false)
        {
            var virtualDir = VirtualPath;
            virtualDir = virtualDir == "/" ? virtualDir : (virtualDir + "/");
            path = path.Replace(ApplicationPath, virtualDir).Replace(@"\", "/");
            return !encode ? path : HttpUtility.UrlEncode(path, System.Text.Encoding.UTF8);
        }

        #region Fields
        /// <summary>
        /// Contains an error message when something goes wrong in the <see cref="ExtractToFolder"/> method.
        /// This message can be retreived with the GetErrorMessage. This way we keep .NET exceptions inside
        /// when this code is called from a COM language
        /// </summary>
        protected string _errorMessage;

        #endregion

        #region internal class
        /// <summary>
        /// Used as a placeholder for the recipients from the MSG file itself or from the "internet"
        /// headers when this message is send outside an Exchange system
        /// </summary>
        private class Recipient
        {
            public string EmailAddress { get; set; }
            public string DisplayName { get; set; }
        }
        #endregion

        #region GetEmailSender
        /// <summary>
        /// Change the E-mail sender addresses to a human readable format
        /// </summary>
        /// <param name="message">The Storage.Message object</param>
        /// <param name="convertToHref">When true the E-mail addresses are converted to hyperlinks</param>
        /// <returns></returns>
        private static string GetEmailSender(Storage.Message message, bool convertToHref = false)
        {
            var output = string.Empty;

            if (message == null) return string.Empty;

            var eMail = message.Sender.Email;
            if (string.IsNullOrEmpty(eMail))
            {
                if (message.Headers != null && message.Headers.From != null)
                    eMail = message.Headers.From.Address;
            }

            var displayName = message.Sender.DisplayName;
            if (string.IsNullOrEmpty(displayName))
            {
                if (message.Headers != null && message.Headers.From != null)
                    displayName = message.Headers.From.DisplayName;
            }

            if (string.IsNullOrEmpty(eMail))
                convertToHref = false;

            if (convertToHref)
                output += "<a href=\"mailto:" + eMail + "\">" +
                          (!string.IsNullOrEmpty(displayName)
                               ? displayName
                               : eMail) + "</a>";

            else
            {
                if (string.IsNullOrEmpty(eMail))
                {
                    output += !string.IsNullOrEmpty(displayName)
                                  ? displayName
                                  : string.Empty;
                }
                else
                {
                    output += eMail +
                              (!string.IsNullOrEmpty(displayName)
                                   ? " (" + displayName + ")"
                                   : string.Empty);
                }
            }

            return output;
        }
        #endregion

        #region GetEmailRecipients
        /// <summary>
        /// Change the E-mail sender addresses to a human readable format
        /// </summary>
        /// <param name="message">The Storage.Message object</param>
        /// <param name="convertToHref">When true the E-mail addresses are converted to hyperlinks</param>
        /// <param name="type">This types says if we want to get the TO's or CC's</param>
        /// <returns></returns>
        private static string GetEmailRecipients(Storage.Message message,
                                                 Storage.Recipient.RecipientType type,
                                                 bool convertToHref = false)
        {
            var output = string.Empty;

            var recipients = new List<Recipient>();

            if (message == null)
                return output;

            foreach (var recipient in message.Recipients)
            {
                // First we filter for the correct recipient type
                if (recipient.Type == type)
                    recipients.Add(new Recipient { EmailAddress = recipient.Email, DisplayName = recipient.DisplayName });
            }

            if (recipients.Count == 0 && message.Headers != null)
            {
                foreach (var to in message.Headers.To)
                    recipients.Add(new Recipient { EmailAddress = to.Address, DisplayName = to.DisplayName });
            }

            foreach (var recipient in recipients)
            {
                if (output != string.Empty)
                    output += "; ";

                var convert = convertToHref;

                if (convert && string.IsNullOrEmpty(recipient.EmailAddress))
                    convert = false;

                if (convert)
                {
                    output += "<a href=\"mailto:" + message.Sender.Email + "\">" +
                              (!string.IsNullOrEmpty(message.Sender.DisplayName)
                                   ? recipient.DisplayName
                                   : recipient.EmailAddress) + "</a>";
                }
                else
                {
                    if (string.IsNullOrEmpty(recipient.EmailAddress))
                    {
                        output += !string.IsNullOrEmpty(recipient.DisplayName)
                                      ? recipient.DisplayName
                                      : string.Empty;
                    }
                    else
                    {
                        output += recipient.EmailAddress +
                                  (!string.IsNullOrEmpty(recipient.DisplayName)
                                       ? " (" + recipient.DisplayName + ")"
                                       : string.Empty);
                    }
                }
            }

            return output;
        }
        #endregion

        #region InjectOutlookHeader
        /// <summary>
        /// Inject an outlook style header into the email body
        /// </summary>
        /// <param name="eMail"></param>
        /// <param name="header"></param>
        /// <returns></returns>
        protected string InjectOutlookHeader(string eMail, string header)
        {
            var temp = eMail.ToUpper();

            var begin = temp.IndexOf("<BODY", StringComparison.Ordinal);

            if (begin > 0)
            {
                begin = temp.IndexOf(">", begin, StringComparison.Ordinal);
                return eMail.Insert(begin + 1, header);
            }

            return header + eMail;
        }
        #endregion

        #region GetInnerException
        /// <summary>
        /// Get the complete inner exception tree
        /// </summary>
        /// <param name="e">The exception object</param>
        /// <returns></returns>
        protected static string GetInnerException(Exception e)
        {
            var exception = e.Message + "\n";
            if (e.InnerException != null)
                exception += GetInnerException(e.InnerException);
            return exception;
        }
        #endregion
    }
}