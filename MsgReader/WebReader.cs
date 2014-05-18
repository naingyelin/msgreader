using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
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

        #region ExtractToFolder
        public string[] ExtractToFolder(string inputFile, string outputFolder, bool hyperlinks = true)
        {
            outputFolder = FileManager.CheckForBackSlash(outputFolder);
            _errorMessage = string.Empty;
            try
            {
                using (var stream = File.Open(inputFile, FileMode.Open, FileAccess.Read))
                using (var message = new Storage.Message(stream))
                {
                    switch (message.Type)
                    {
                        case Storage.Message.MessageType.Email:
                            return WriteEmail(message, outputFolder, hyperlinks).ToArray();

                        case Storage.Message.MessageType.AppointmentRequest:
                        case Storage.Message.MessageType.Appointment:
                        case Storage.Message.MessageType.AppointmentResponse:
                            return WriteAppointment(message, outputFolder, hyperlinks).ToArray();
                        case Storage.Message.MessageType.Task:
                            throw new Exception("A task file is not supported");
                        case Storage.Message.MessageType.StickyNote:
                            return WriteStickyNote(message, outputFolder, hyperlinks).ToArray();
                        case Storage.Message.MessageType.Unknown:
                            try
                            {
                                return WriteEmail(message, outputFolder, hyperlinks).ToArray();
                            }
                            catch
                            {
                                throw new NotSupportedException("Unknown message type");
                            }
                    }
                }
            }
            catch (Exception e)
            {
                _errorMessage = GetInnerException(e);
                return new string[0];
            }

            // If we return here then the file was not supported
            return new string[0];
        }
        #endregion

        #region WriteEmail
        private List<string> WriteEmail(Storage.Message message, string outputFolder, bool hyperlinks)
        {
            var result = new List<string>();

            // Read MSG file from a stream
            // We first always check if there is a HTML body
            var body = message.BodyHtml;
            var htmlBody = true;

            if (body == null)
            {
                // When there is not HTML body found then try to get the text body
                body = message.BodyText;
                htmlBody = false;
            }

            // Determine the name for the E-mail body
            var eMailFileName = outputFolder + Path.GetFileNameWithoutExtension(FileManager.RemoveInvalidFileNameChars(message.FileName)) + ".html";
            //(!string.IsNullOrEmpty(message.FileName)
            //     ? FileManager.RemoveInvalidFileNameChars(message.FileName)
            //     : "email") + (htmlBody ? ".html" : ".txt");

            result.Add(eMailFileName);

            var attachmentList = new List<string>();

            foreach (var attachment in message.Attachments)
            {
                FileInfo fileInfo = null;

                if (attachment is Storage.Attachment)
                {
                    var attach = (Storage.Attachment)attachment;
                    var fileName = FileManager.FileExistsMakeNew(outputFolder +
                                                  FileManager.RemoveInvalidFileNameChars(attach.FileName)
                                                             .Replace("&", "[and]"));
                    fileInfo = new FileInfo(FileManager.FileExistsMakeNew(fileName));
                    File.WriteAllBytes(fileInfo.FullName, attach.Data);

                    // When we find an in-line attachment we have to replace the CID tag inside the HTML body
                    // with the name of the in-line attachment. But before we do this we check if the CID exists.
                    // When the CID does not exists we treat the in-line attachment as a normal attachment
                    if (htmlBody && !string.IsNullOrEmpty(attach.ContentId) &&
                        body.Contains(attach.ContentId))
                    {
                        body = body.Replace("cid:" + attach.ContentId,
                                            GetRelativePathFromAbsolutePath(fileInfo.FullName));
                        continue;
                    }

                    result.Add(fileInfo.FullName);
                }
                else if (attachment is Storage.Message)
                {
                    var msg = (Storage.Message)attachment;
                    var fileName = FileManager.FileExistsMakeNew(outputFolder +
                        FileManager.RemoveInvalidFileNameChars(msg.Subject).Replace("&", "[and]") + ".msg");
                    fileInfo = new FileInfo(fileName);
                    result.Add(fileInfo.FullName);
                    msg.Save(fileInfo.FullName);
                }

                if (fileInfo == null) continue;

                attachmentList.Add(BuildAnchor(fileInfo.FullName, MsgHandlerUrl));
            }

            string emailHeader;

            if (htmlBody)
            {
                // Start of table
                emailHeader =
                    "<table style=\"width:100%; font-family: Times New Roman; font-size: 12pt;\">" + Environment.NewLine;

                // From
                emailHeader +=
                    "<tr style=\"height: 18px; vertical-align: top; \"><td style=\"width: 100px; font-weight: bold; \">" + LanguageConsts.EmailFromLabel + ":</td><td>" + GetEmailSender(message, hyperlinks, true) + "</td></tr>" + Environment.NewLine;

                // Sent on
                if (message.SentOn != null)
                    emailHeader +=
                        "<tr style=\"height: 18px; vertical-align: top; \"><td style=\"width: 100px; font-weight: bold; \">" + LanguageConsts.EmailSentOnLabel + ":</td><td>" + ((DateTime)message.SentOn).ToString(LanguageConsts.DataFormat) + "</td></tr>" + Environment.NewLine;

                // To
                emailHeader +=
                    "<tr style=\"height: 18px; vertical-align: top; \"><td style=\"width: 100px; font-weight: bold; \">" +
                    LanguageConsts.EmailToLabel + ":</td><td>" +
                    GetEmailRecipients(message, Storage.Recipient.RecipientType.To, hyperlinks, true) + "</td></tr>" +
                    Environment.NewLine;

                // CC
                var cc = GetEmailRecipients(message, Storage.Recipient.RecipientType.Cc, hyperlinks, false);
                if (cc != string.Empty)
                    emailHeader +=
                        "<tr style=\"height: 18px; vertical-align: top; \"><td style=\"width: 100px; font-weight: bold; \">" +
                        LanguageConsts.EmailCcLabel + ":</td><td>" + cc + "</td></tr>" + Environment.NewLine;

                // BCC
                var bcc = GetEmailRecipients(message, Storage.Recipient.RecipientType.Bcc, hyperlinks, false);
                if (bcc != string.Empty)
                    emailHeader +=
                        "<tr style=\"height: 18px; vertical-align: top; \"><td style=\"width: 100px; font-weight: bold; \">" +
                        LanguageConsts.EmailBccLabel + ":</td><td>" + bcc + "</td></tr>" + Environment.NewLine;

                // Subject
                emailHeader +=
                    "<tr style=\"height: 18px; vertical-align: top; \"><td style=\"width: 100px; font-weight: bold; \">" +
                    LanguageConsts.EmailSubjectLabel + ":</td><td>" + message.Subject + "</td></tr>" + Environment.NewLine;

                // Urgent
                if (message.Importance != Storage.Message.MessageImportance.Normal)
                {
                    var importanceText = LanguageConsts.ImportanceLowText;
                    if (message.Importance == Storage.Message.MessageImportance.High)
                        importanceText = LanguageConsts.ImportanceHighText;

                    emailHeader +=
                        "<tr style=\"height: 18px; vertical-align: top; \"><td style=\"width: 100px; font-weight: bold; \">" +
                        LanguageConsts.ImportanceLabel + ":</td><td>" + importanceText + "</td></tr>" + Environment.NewLine;

                    // Empty line
                    emailHeader += "<tr><td colspan=\"2\" style=\"height: 18px; \">&nbsp</td></tr>" + Environment.NewLine;
                }

                // Attachments
                if (attachmentList.Count != 0)
                    emailHeader +=
                        "<tr style=\"height: 18px; vertical-align: top; \"><td style=\"width: 100px; font-weight: bold; \">" +
                        LanguageConsts.EmailAttachmentsLabel + ":</td><td>" + string.Join(", ", attachmentList.ToArray()) + "</td></tr>" +
                        Environment.NewLine;

                // Empty line
                emailHeader += "<tr><td colspan=\"2\" style=\"height: 18px; \">&nbsp</td></tr>" + Environment.NewLine;

                // Follow up
                if (message.Flag != null)
                {
                    emailHeader +=
                        "<tr style=\"height: 18px; vertical-align: top; \"><td style=\"width: 100px; font-weight: bold; \">" +
                        LanguageConsts.EmailFollowUpLabel + ":</td><td>" + message.Flag.Request + "</td></tr>" + Environment.NewLine;

                    // When complete
                    if (message.Task.Complete != null && (bool)message.Task.Complete)
                    {
                        emailHeader +=
                            "<tr style=\"height: 18px; vertical-align: top; \"><td style=\"width: 100px; font-weight: bold; \">" +
                            LanguageConsts.EmailFollowUpStatusLabel + ":</td><td>" + LanguageConsts.EmailFollowUpCompletedText +
                            "</td></tr>" + Environment.NewLine;

                        // Task completed date
                        var completedDate = message.Task.CompleteTime;
                        if (completedDate != null)
                            emailHeader +=
                                "<tr style=\"height: 18px; vertical-align: top; \"><td style=\"width: 100px; font-weight: bold; \">" +
                                LanguageConsts.EmailTaskDateCompleted + ":</td><td>" + completedDate + "</td></tr>" + Environment.NewLine;
                    }
                    else
                    {
                        // Task startdate
                        var startDate = message.Task.StartDate;
                        if (startDate != null)
                            emailHeader +=
                                "<tr style=\"height: 18px; vertical-align: top; \"><td style=\"width: 100px; font-weight: bold; \">" +
                                LanguageConsts.EmailTaskStartDateLabel + ":</td><td>" + startDate + "</td></tr>" + Environment.NewLine;

                        // Task duedate
                        var dueDate = message.Task.DueDate;
                        if (dueDate != null)
                            emailHeader +=
                                "<tr style=\"height: 18px; vertical-align: top; \"><td style=\"width: 100px; font-weight: bold; \">" +
                                LanguageConsts.EmailTaskDueDateLabel + ":</td><td>" + dueDate + "</td></tr>" + Environment.NewLine;

                    }

                    // Empty line
                    emailHeader += "<tr><td colspan=\"2\" style=\"height: 18px; \">&nbsp</td></tr>" + Environment.NewLine;
                }

                // Categories
                var categories = message.Categories;
                if (categories != null)
                {
                    emailHeader +=
                        "<tr style=\"height: 18px; vertical-align: top; \"><td style=\"width: 100px; font-weight: bold; \">" +
                        LanguageConsts.EmailCategoriesLabel + ":</td><td>" + String.Join("; ", categories.ToArray()) + "</td></tr>" +
                        Environment.NewLine;

                    // Empty line
                    emailHeader += "<tr><td colspan=\"2\" style=\"height: 18px; \">&nbsp</td></tr>" + Environment.NewLine;
                }

                // End of table + empty line
                emailHeader += "</table><br/>" + Environment.NewLine;

                body = InjectHeader(body, emailHeader);
            }
            else
            {
                // Read all the language consts and get the longest string
                var languageConsts = new List<string>
                {
                    LanguageConsts.EmailFromLabel,
                    LanguageConsts.EmailSentOnLabel,
                    LanguageConsts.EmailToLabel,
                    LanguageConsts.EmailCcLabel,
                    LanguageConsts.EmailBccLabel,
                    LanguageConsts.EmailSubjectLabel,
                    LanguageConsts.ImportanceLabel,
                    LanguageConsts.EmailAttachmentsLabel,
                    LanguageConsts.EmailFollowUpFlag,
                    LanguageConsts.EmailFollowUpLabel,
                    LanguageConsts.EmailFollowUpStatusLabel,
                    LanguageConsts.EmailFollowUpCompletedText,
                    LanguageConsts.EmailTaskStartDateLabel,
                    LanguageConsts.EmailTaskDueDateLabel,
                    LanguageConsts.EmailTaskDateCompleted,
                    LanguageConsts.EmailCategoriesLabel
                };

                var maxLength = languageConsts.Select(languageConst => languageConst.Length).Concat(new[] { 0 }).Max();

                // From
                emailHeader =
                    (LanguageConsts.EmailFromLabel + ":").PadRight(maxLength) + GetEmailSender(message, false, false) + Environment.NewLine;

                // Sent on
                if (message.SentOn != null)
                    emailHeader +=
                        (LanguageConsts.EmailSentOnLabel + ":").PadRight(maxLength) +
                        ((DateTime)message.SentOn).ToString(LanguageConsts.DataFormat) + Environment.NewLine;

                // To
                emailHeader +=
                    (LanguageConsts.EmailToLabel + ":").PadRight(maxLength) +
                    GetEmailRecipients(message, Storage.Recipient.RecipientType.To, false, false) + Environment.NewLine;

                // CC
                var cc = GetEmailRecipients(message, Storage.Recipient.RecipientType.Cc, false, false);
                if (cc != string.Empty)
                    emailHeader += (LanguageConsts.EmailCcLabel + ":").PadRight(maxLength) + cc + Environment.NewLine;

                // CC
                var bcc = GetEmailRecipients(message, Storage.Recipient.RecipientType.Bcc, false, false);
                if (bcc != string.Empty)
                    emailHeader += (LanguageConsts.EmailCcLabel + ":").PadRight(maxLength) + bcc + Environment.NewLine;

                // Subject
                emailHeader += (LanguageConsts.EmailSubjectLabel + ":").PadRight(maxLength) + message.Subject + Environment.NewLine;

                if (message.Importance != Storage.Message.MessageImportance.Normal)
                {
                    var importanceText = LanguageConsts.ImportanceLowText;
                    if (message.Importance == Storage.Message.MessageImportance.High)
                        importanceText = LanguageConsts.ImportanceHighText;

                    // Importance text + new line
                    emailHeader += (LanguageConsts.ImportanceLabel + ":").PadRight(maxLength) + importanceText + Environment.NewLine + Environment.NewLine;
                }

                // Attachments
                if (attachmentList.Count != 0)
                    emailHeader += (LanguageConsts.EmailAttachmentsLabel + ":").PadRight(maxLength) +
                                          string.Join(", ", attachmentList.ToArray()) + Environment.NewLine + Environment.NewLine;

                // Urgent
                if (message.Importance != Storage.Message.MessageImportance.Normal)
                {
                    var importanceText = LanguageConsts.ImportanceLowText;
                    if (message.Importance == Storage.Message.MessageImportance.High)
                        importanceText = LanguageConsts.ImportanceHighText;

                    emailHeader += (LanguageConsts.ImportanceLabel + ":").PadRight(maxLength) +
                                    importanceText + Environment.NewLine + Environment.NewLine;
                }

                // Follow up
                if (message.Flag != null)
                {
                    emailHeader += (LanguageConsts.EmailFollowUpLabel + ":").PadRight(maxLength) + message.Flag.Request + Environment.NewLine;

                    // When complete
                    if (message.Task.Complete != null && (bool)message.Task.Complete)
                    {
                        emailHeader += (LanguageConsts.EmailFollowUpStatusLabel + ":").PadRight(maxLength) +
                                              LanguageConsts.EmailFollowUpCompletedText + Environment.NewLine;

                        // Task completed date
                        var completedDate = message.Task.CompleteTime;
                        if (completedDate != null)
                            emailHeader += (LanguageConsts.EmailTaskDateCompleted + ":").PadRight(maxLength) + completedDate + Environment.NewLine;
                    }
                    else
                    {
                        // Task startdate
                        var startDate = message.Task.StartDate;
                        if (startDate != null)
                            emailHeader += (LanguageConsts.EmailTaskStartDateLabel + ":").PadRight(maxLength) + startDate + Environment.NewLine;

                        // Task duedate
                        var dueDate = message.Task.DueDate;
                        if (dueDate != null)
                            emailHeader += (LanguageConsts.EmailTaskDueDateLabel + ":").PadRight(maxLength) + dueDate + Environment.NewLine;

                    }

                    // Empty line
                    emailHeader += Environment.NewLine;
                }

                // Categories
                var categories = message.Categories;
                if (categories != null)
                {
                    emailHeader += (LanguageConsts.EmailCategoriesLabel + ":").PadRight(maxLength) +
                                          String.Join("; ", categories.ToArray()) + Environment.NewLine;

                    // Empty line
                    emailHeader += Environment.NewLine;
                }


                body = emailHeader + body;
            }

            // Write the body to a file
            File.WriteAllText(eMailFileName, body, Encoding.UTF8);

            return result;
        }
        #endregion

        #region WriteStickyNote
        /// <summary>
        /// Writes the body of the MSG StickyNote to html or text and extracts all the attachments. The
        /// result is return as a List of strings
        /// </summary>
        /// <param name="message"><see cref="Storage.Message"/></param>
        /// <param name="outputFolder">The folder where we need to write the output</param>
        /// <param name="hyperlinks">When true then hyperlinks are generated for the To, CC, BCC and attachments</param>
        /// <returns></returns>
        private List<string> WriteStickyNote(Storage.Message message, string outputFolder, bool hyperlinks)
        {
            var result = new List<string>();
            string stickyNoteFile;
            var stickyNoteHeader = string.Empty;

            // Sticky notes only have RTF or Text bodies
            var body = message.BodyRtf;

            // If the body is not null then we convert it to HTML
            if (body != null)
            {
                var converter = new RtfToHtmlConverter();
                body = converter.ConvertRtfToHtml(body);
                stickyNoteFile = outputFolder +
                                 (!string.IsNullOrEmpty(message.Subject)
                                     ? FileManager.RemoveInvalidFileNameChars(message.Subject)
                                     : "stickynote") + ".htm";

                stickyNoteHeader =
                    "<table style=\"width:100%; font-family: Times New Roman; font-size: 12pt;\">" + Environment.NewLine;

                if (message.SentOn != null)
                    stickyNoteHeader +=
                        "<tr style=\"height: 18px; vertical-align: top; \"><td style=\"width: 100px; font-weight: bold; \">" +
                        LanguageConsts.StickyNoteDateLabel + ":</td><td>" +
                        ((DateTime)message.SentOn).ToString(LanguageConsts.DataFormat) + "</td></tr>" +
                        Environment.NewLine;

                // Empty line
                stickyNoteHeader += "<tr><td colspan=\"2\" style=\"height: 18px; \">&nbsp</td></tr>" + Environment.NewLine;

                // End of table + empty line
                stickyNoteHeader += "</table><br/>" + Environment.NewLine;

                body = InjectHeader(body, stickyNoteHeader);
            }
            else
            {
                body = message.BodyText ?? string.Empty;

                // Sent on
                if (message.SentOn != null)
                    stickyNoteHeader +=
                        (LanguageConsts.StickyNoteDateLabel + ":") + ((DateTime)message.SentOn).ToString(LanguageConsts.DataFormat) + Environment.NewLine;

                body = stickyNoteHeader + body;
                stickyNoteFile = outputFolder + (!string.IsNullOrEmpty(message.Subject) ? FileManager.RemoveInvalidFileNameChars(message.Subject) : "stickynote") + ".txt";
            }

            // Write the body to a file
            File.WriteAllText(stickyNoteFile, body, Encoding.UTF8);
            result.Add(stickyNoteFile);
            return result;
        }
        #endregion

        #region WriteAppointment
        /// <summary>
        /// Writes the body of the MSG Appointment to html or text and extracts all the attachments. The
        /// result is return as a List of strings
        /// </summary>
        /// <param name="message"><see cref="Storage.Message"/></param>
        /// <param name="outputFolder">The folder where we need to write the output</param>
        /// <param name="hyperlinks">When true then hyperlinks are generated for the To, CC, BCC and attachments</param>
        /// <returns></returns>
        private List<string> WriteAppointment(Storage.Message message, string outputFolder, bool hyperlinks)
        {
            //throw new NotImplementedException("Todo");
            // TODO: Rewrite this code so that an correct appointment is written

            var result = new List<string>();

            // Read MSG file from a stream
            // We first always check if there is a RTF body because appointments never have HTML bodies
            var body = message.BodyRtf;
            var htmlBody = false;

            // If the body is not null then we convert it to HTML
            if (body != null)
            {
                var converter = new RtfToHtmlConverter();
                body = converter.ConvertRtfToHtml(body);
                htmlBody = true;
            }

            if (string.IsNullOrEmpty(body))
            {
                body = message.BodyText;
                if (body == null)
                {
                    body = "<html><head></head><body></body></html>";
                    htmlBody = true;
                }
            }

            // Determine the name for the appointment body
            var appointmentFileName = outputFolder +
                                      (!string.IsNullOrEmpty(message.Subject)
                                          ? FileManager.RemoveInvalidFileNameChars(message.Subject)
                                          : "appointment") + (htmlBody ? ".htm" : ".txt");

            result.Add(appointmentFileName);

            // Onderwerp
            // Locatie
            //
            // Begin
            // Eind
            // Tijd weergeven als
            //
            // Terugkeerpatroon
            // Type terugkeerpatroon
            //
            // Vergaderingstatus
            //
            // Organisator
            // Verplichte deelnemers
            // Optionele deelnemers
            // 
            // Inhoud van het agenda item

            string appointmentHeader;

            if (htmlBody)
            {
                // Start of table
                appointmentHeader =
                    "<table style=\"width:100%; font-family: Times New Roman; font-size: 12pt;\">" + Environment.NewLine;

                // Subject
                appointmentHeader +=
                    "<tr style=\"height: 18px; vertical-align: top; \"><td style=\"width: 100px; font-weight: bold; \">" +
                    LanguageConsts.AppointmentSubject + ":</td><td>" + message.Subject + "</td></tr>" + Environment.NewLine;

                // Location
                appointmentHeader +=
                    "<tr style=\"height: 18px; vertical-align: top; \"><td style=\"width: 100px; font-weight: bold; \">" +
                    LanguageConsts.AppointmentLocation + ":</td><td>" + message.Appointment.Location + "</td></tr>" + Environment.NewLine;

                // Empty line
                appointmentHeader += "<tr><td colspan=\"2\" style=\"height: 18px; \">&nbsp</td></tr>" + Environment.NewLine;

                // Start
                appointmentHeader +=
                    "<tr style=\"height: 18px; vertical-align: top; \"><td style=\"width: 100px; font-weight: bold; \">" +
                    LanguageConsts.AppointmentStartDate + ":</td><td>" + message.Appointment.Start + "</td></tr>" + Environment.NewLine;

                // End
                appointmentHeader +=
                    "<tr style=\"height: 18px; vertical-align: top; \"><td style=\"width: 100px; font-weight: bold; \">" +
                    LanguageConsts.AppointmentEndDate + ":</td><td>" + message.Appointment.End + "</td></tr>" + Environment.NewLine;

                // Empty line
                appointmentHeader += "<tr><td colspan=\"2\" style=\"height: 18px; \">&nbsp</td></tr>" + Environment.NewLine;

                // Recurrence patern
                var recurrenceType = message.Appointment.RecurrenceType;
                if (!string.IsNullOrEmpty(recurrenceType))
                    appointmentHeader +=
                        "<tr style=\"height: 18px; vertical-align: top; \"><td style=\"width: 100px; font-weight: bold; \">" +
                        LanguageConsts.AppointmentRecurrenceTypeLabel + ":</td><td>" + message.Appointment.RecurrenceType + "</td></tr>" + Environment.NewLine;

                // Recurrence patern
                var recurrencePatern = message.Appointment.RecurrencePatern;
                if (!string.IsNullOrEmpty(recurrencePatern))
                    appointmentHeader +=
                        "<tr style=\"height: 18px; vertical-align: top; \"><td style=\"width: 100px; font-weight: bold; \">" +
                        LanguageConsts.AppointmentRecurrencePaternLabel + ":</td><td>" + message.Appointment.RecurrencePatern + "</td></tr>" + Environment.NewLine;


                // Categories
                var categories = message.Categories;
                if (categories != null)
                {
                    appointmentHeader +=
                        "<tr style=\"height: 18px; vertical-align: top; \"><td style=\"width: 100px; font-weight: bold; \">" +
                        LanguageConsts.EmailCategoriesLabel + ":</td><td>" + String.Join("; ", categories.ToArray()) + "</td></tr>" +
                        Environment.NewLine;

                    // Empty line
                    appointmentHeader += "<tr><td colspan=\"2\" style=\"height: 18px; \">&nbsp</td></tr>" + Environment.NewLine;
                }

                // Attachments
                //if (attachmentList.Count != 0)
                //    appointmentHeader +=
                //        "<tr style=\"height: 18px; vertical-align: top; \"><td style=\"width: 100px; font-weight: bold; \">" +
                //        LanguageConsts.AttachmentsLabel + ":</td><td>" + string.Join(", ", attachmentList) + "</td></tr>" +
                //        Environment.NewLine;

                // Empty line
                appointmentHeader += "<tr><td colspan=\"2\" style=\"height: 18px; \">&nbsp</td></tr>" + Environment.NewLine;

                // End of table + empty line
                appointmentHeader += "</table><br/>" + Environment.NewLine;
            }
            else
            {
                // text part, todo
                appointmentHeader = "todo";
            }

            body = InjectHeader(body, appointmentHeader);


            // Write the body to a file
            File.WriteAllText(appointmentFileName, body, Encoding.UTF8);

            return result;
        }
        #endregion

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
        /// <param name="convertToHref">When true then E-mail addresses are converted to hyper-links</param>
        /// <param name="html">Set this to true when the E-mail body format is HTML</param>
        /// <returns></returns>
        private static string GetEmailSender(Storage.Message message, bool convertToHref, bool html)
        {
            var output = string.Empty;

            if (message == null) return string.Empty;

            var tempEmailAddress = message.Sender.Email;
            var tempDisplayName = message.Sender.DisplayName;

            if (string.IsNullOrEmpty(tempEmailAddress) && message.Headers != null && message.Headers.From != null)
                tempEmailAddress = RemoveSingleQuotes(message.Headers.From.Address);

            if (string.IsNullOrEmpty(tempDisplayName) && message.Headers != null && message.Headers.From != null)
                tempDisplayName = message.Headers.From.DisplayName;

            var emailAddress = tempEmailAddress;
            var displayName = tempDisplayName;

            // Sometimes the E-mail address and displayname get swapped so check if they are valid
            if (!IsEmailAddressValid(tempEmailAddress) && IsEmailAddressValid(tempDisplayName))
            {
                // Swap them
                emailAddress = tempDisplayName;
                displayName = tempEmailAddress;
            }
            else if (IsEmailAddressValid(tempDisplayName))
            {
                // If the displayname is an emailAddress them move it
                emailAddress = tempDisplayName;
                displayName = tempDisplayName;
            }

            if (html)
            {
                emailAddress = HttpUtility.HtmlEncode(emailAddress);
                displayName = HttpUtility.HtmlEncode(displayName);
            }

            if (convertToHref && html && !string.IsNullOrEmpty(emailAddress))
                output += "<a href=\"mailto:" + emailAddress + "\">" +
                          (!string.IsNullOrEmpty(displayName)
                              ? displayName
                              : emailAddress) + "</a>";

            else
            {
                if (!string.IsNullOrEmpty(emailAddress))
                    output = emailAddress;

                if (!string.IsNullOrEmpty(displayName))
                    output += (!string.IsNullOrEmpty(emailAddress) ? " <" : string.Empty) + displayName +
                              (!string.IsNullOrEmpty(emailAddress) ? ">" : string.Empty);
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
        //private static string GetEmailRecipients(Storage.Message message,
        //                                         Storage.Recipient.RecipientType type,
        //                                         bool convertToHref = false)
        //{
        //    var output = string.Empty;

        //    var recipients = new List<Recipient>();

        //    if (message == null)
        //        return output;

        //    foreach (var recipient in message.Recipients)
        //    {
        //        // First we filter for the correct recipient type
        //        if (recipient.Type == type)
        //            recipients.Add(new Recipient { EmailAddress = recipient.Email, DisplayName = recipient.DisplayName });
        //    }

        //    if (recipients.Count == 0 && message.Headers != null)
        //    {
        //        foreach (var to in message.Headers.To)
        //            recipients.Add(new Recipient { EmailAddress = to.Address, DisplayName = to.DisplayName });
        //    }

        //    foreach (var recipient in recipients)
        //    {
        //        if (output != string.Empty)
        //            output += "; ";

        //        var convert = convertToHref;

        //        if (convert && string.IsNullOrEmpty(recipient.EmailAddress))
        //            convert = false;

        //        if (convert)
        //        {
        //            output += "<a href=\"mailto:" + message.Sender.Email + "\">" +
        //                      (!string.IsNullOrEmpty(message.Sender.DisplayName)
        //                           ? recipient.DisplayName
        //                           : recipient.EmailAddress) + "</a>";
        //        }
        //        else
        //        {
        //            if (string.IsNullOrEmpty(recipient.EmailAddress))
        //            {
        //                output += !string.IsNullOrEmpty(recipient.DisplayName)
        //                              ? recipient.DisplayName
        //                              : string.Empty;
        //            }
        //            else
        //            {
        //                output += recipient.EmailAddress +
        //                          (!string.IsNullOrEmpty(recipient.DisplayName)
        //                               ? " (" + recipient.DisplayName + ")"
        //                               : string.Empty);
        //            }
        //        }
        //    }

        //    return output;
        //}
        #endregion

        #region GetEmailRecipients
        /// <summary>
        /// Change the E-mail sender addresses to a human readable format
        /// </summary>
        /// <param name="message">The Storage.Message object</param>
        /// <param name="convertToHref">When true the E-mail addresses are converted to hyperlinks</param>
        /// <param name="type">This types says if we want to get the TO's or CC's</param>
        /// <param name="html">Set this to true when the E-mail body format is html</param>
        /// <returns></returns>
        private static string GetEmailRecipients(Storage.Message message,
                                                 Storage.Recipient.RecipientType type,
                                                 bool convertToHref,
                                                 bool html)
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
                switch (type)
                {
                    case Storage.Recipient.RecipientType.To:
                        if (message.Headers.To != null)
                            recipients.AddRange(message.Headers.To.Select(to => new Recipient { EmailAddress = to.Address, DisplayName = to.DisplayName }));
                        break;

                    case Storage.Recipient.RecipientType.Cc:
                        if (message.Headers.Cc != null)
                            recipients.AddRange(message.Headers.Cc.Select(cc => new Recipient { EmailAddress = cc.Address, DisplayName = cc.DisplayName }));
                        break;

                    case Storage.Recipient.RecipientType.Bcc:
                        if (message.Headers.Bcc != null)
                            recipients.AddRange(message.Headers.Bcc.Select(bcc => new Recipient { EmailAddress = bcc.Address, DisplayName = bcc.DisplayName }));
                        break;
                }
            }

            foreach (var recipient in recipients)
            {
                if (output != string.Empty)
                    output += "; ";

                var tempEmailAddress = RemoveSingleQuotes(recipient.EmailAddress);
                var tempDisplayName = RemoveSingleQuotes(recipient.DisplayName);

                var emailAddress = tempEmailAddress;
                var displayName = tempDisplayName;

                // Sometimes the E-mail address and displayname get swapped so check if they are valid
                if (!IsEmailAddressValid(tempEmailAddress) && IsEmailAddressValid(tempDisplayName))
                {
                    // Swap them
                    emailAddress = tempDisplayName;
                    displayName = tempEmailAddress;
                }
                else if (IsEmailAddressValid(tempDisplayName))
                {
                    // If the displayname is an emailAddress them move it
                    emailAddress = tempDisplayName;
                    displayName = tempDisplayName;
                }

                if (html)
                {
                    emailAddress = HttpUtility.HtmlEncode(emailAddress);
                    displayName = HttpUtility.HtmlEncode(displayName);
                }

                if (convertToHref && html && !string.IsNullOrEmpty(emailAddress))
                    output += "<a href=\"mailto:" + emailAddress + "\">" +
                              (!string.IsNullOrEmpty(displayName)
                                  ? displayName
                                  : emailAddress) + "</a>";

                else
                {
                    if (!string.IsNullOrEmpty(emailAddress))
                        output = emailAddress;

                    if (!string.IsNullOrEmpty(displayName))
                        output += (!string.IsNullOrEmpty(emailAddress) ? " <" : string.Empty) + displayName +
                                  (!string.IsNullOrEmpty(emailAddress) ? ">" : string.Empty);
                }
            }

            return output;
        }
        #endregion

        #region IsEmailAddressValid
        /// <summary>
        /// Return true when the E-mail address is valid
        /// </summary>
        /// <param name="emailAddress"></param>
        /// <returns></returns>
        private static bool IsEmailAddressValid(string emailAddress)
        {
            if (string.IsNullOrEmpty(emailAddress))
                return false;

            var regex = new Regex(@"\w+([-+.]\w+)*@\w+([-.]\w+)*\.\w+([-.]\w+)*", RegexOptions.IgnoreCase);
            var matches = regex.Matches(emailAddress);

            return matches.Count == 1;
        }
        #endregion

        #region RemoveSingleQuotes
        /// <summary>
        /// Removes trailing en ending single quotes from an E-mail address when they exist
        /// </summary>
        /// <param name="email"></param>
        /// <returns></returns>
        private static string RemoveSingleQuotes(string email)
        {
            if (string.IsNullOrEmpty(email))
                return string.Empty;

            if (email.StartsWith("'"))
                email = email.Substring(1, email.Length - 1);

            if (email.EndsWith("'"))
                email = email.Substring(0, email.Length - 1);

            return email;
        }
        #endregion

        #region InjectHeader
        /// <summary>
        /// Inject an outlook style header into the email body
        /// </summary>
        /// <param name="eMail"></param>
        /// <param name="header"></param>
        /// <returns></returns>
        protected string InjectHeader(string body, string header)
        {
            var temp = body.ToUpperInvariant();

            var begin = temp.IndexOf("<BODY", StringComparison.Ordinal);

            if (begin > 0)
            {
                begin = temp.IndexOf(">", begin, StringComparison.Ordinal);
                return body.Insert(begin + 1, header);
            }

            return header + body;
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