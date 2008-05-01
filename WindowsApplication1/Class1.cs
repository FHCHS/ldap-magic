//<add key="Email_SMTP_Server" value=""/>
//<add key="Email_Error_From" value=""/>
//<add key="Email_Error_FromName" value=""/>
//<add key="Email_Error_Recipient" value=""/>
//<add key="Email_Error_Subject" value=""/>
//<add key="Email_SMTP_UserName" value=""/>
//<add key="Email_SMTP_Password" value=""/>

using System;
using System.Configuration;
using System.Net.Mail;
using System.Web;
using System.Collections.Specialized;
using System.Collections.Generic;
using System.Text;

namespace WindowsApplication1
{


	/// <summary>
	/// Class to send email.
	/// </summary>
	public static class Email
	{


		/// <summary>
		/// Helper function for sending Error emails.
		/// </summary>
		/// <param name="message">The email message.</param>
        public static void SendErrorEmail(string message)
		{
			System.Configuration.AppSettingsReader asr = new AppSettingsReader();

			SendEmail(
				Convert.ToString(asr.GetValue("Email_SMTP_Server",Type.GetType("System.String"))),
				Convert.ToString(asr.GetValue("Email_Error_From",Type.GetType("System.String"))),
				Convert.ToString(asr.GetValue("Email_Error_FromName",Type.GetType("System.String"))),
				Convert.ToString(asr.GetValue("Email_Error_Recipient",Type.GetType("System.String"))),
				Convert.ToString(asr.GetValue("Email_Error_Subject",Type.GetType("System.String"))),
				message);

		}

		/// <summary>
		/// Send the email.
		/// </summary>
		/// <param name="SMTPServer">The name or IP of the SMTP Email server to use. An empty string indicates the default SMTP server on this server should be used.</param>
		/// <param name="emailFrom">Who the email is from.</param>
		/// <param name="emailFromName">The name of the sender</param>
		/// <param name="emailTo">Who the email is to.</param>
		/// <param name="subject">The Subject of the email.</param>
		/// <param name="message">The email message</param>
		public static void SendEmail(string SMTPServer, string emailFrom, string emailFromName, string emailTo, string subject,
			string message)
		{
            if (emailTo != null)
            {
                // Split multiple mail receipients
                string[] straEmailToList;		// Array of the email recipietns
                straEmailToList = emailTo.Split(';');

                // Determine the Sender. Include the email and name if present.
                MailAddress maFromAddress;
                if (emailFromName != "")
                    maFromAddress = new MailAddress(emailFrom, emailFromName);
                else
                    maFromAddress = new MailAddress(emailFrom);

                // Create the Mail Message
                MailMessage mmMsg = new MailMessage();
                mmMsg.From = maFromAddress;
                foreach (string strEmailTo in straEmailToList)
                {
                    mmMsg.To.Add(new MailAddress(strEmailTo));
                }
                mmMsg.Subject = subject;
                mmMsg.Body = message;

                // Send the email
                try
                {
                    SmtpClient smtpcMessage = new SmtpClient(SMTPServer);

                    if (((Convert.ToString(ConfigurationSettings.AppSettings["Email_SMTP_UserName"]) + "") != "") && ((Convert.ToString(ConfigurationSettings.AppSettings["Email_SMTP_Password"]) + "") != ""))
                    {
                        smtpcMessage.Credentials = new System.Net.NetworkCredential(ConfigurationSettings.AppSettings["Email_SMTP_UserName"], ConfigurationSettings.AppSettings["Email_SMTP_Password"]);
                    }
                    
                     smtpcMessage.Send(mmMsg);
                }
                catch (Exception ex)
                {
                    // Do nothing if there was an error as there is nothing we can do. 
                    return;
                }
            }
		}

		/// <summary>
		/// Send an email with an attachment.
		/// </summary>
		/// <param name="SMTPServer">The naame or IP of the SMTP Email server to use. An empty string indicates the default SMTP server on this server should be used.</param>
		/// <param name="emailFrom">Who the email is from.</param>
		/// <param name="emailFromName">The name of the sender</param>
		/// <param name="emailTo">Who the email is to.</param>
		/// <param name="subject">The Subject of the email.</param>
		/// <param name="message">The email message</param>
		/// <param name="attachmentPath">The complete path to the attachment on the web server including the file name. The path should be the local path instead of the website path, i.e. c:\inetpub\uploads\info.doc</param>
		public static void SendEmailWithAttachment(string SMTPServer, string emailFrom, string emailFromName, string emailTo,
			string subject, string message, string attachmentPath)
		{
			string strEmailFrom = "";		// The complete sender for the email (including the name if present)

			// Determine the Sender. Include the email and name if present.
			if (emailFromName != "")
				strEmailFrom = emailFromName + " <" + emailFrom + ">";
			else
				strEmailFrom = emailFrom;

			// Send the email
			try
			{
				SmtpClient smtpcMessage = new SmtpClient(SMTPServer);

				MailMessage mmMessage = new MailMessage(strEmailFrom, emailTo, subject, message);

				if (attachmentPath.Trim() != "")
				{
					Attachment aAttachment = new Attachment(attachmentPath);
					mmMessage.Attachments.Add(aAttachment);
				}

				smtpcMessage.Send(mmMessage);
			}
			catch
			{
				// Do nothing if there was an error as there is nothing we can do. 
				return;
			}

		}

        /// <summary>
        /// Send an email with an attachment.
        /// </summary>
        /// <param name="SMTPServer">The naame or IP of the SMTP Email server to use. An empty string indicates the default SMTP server on this server should be used.</param>
        /// <param name="emailFrom">Who the email is from.</param>
        /// <param name="emailFromName">The name of the sender</param>
        /// <param name="emailTo">Who the email is to.</param>
        /// <param name="subject">The Subject of the email.</param>
        /// <param name="message">The email message</param>
        /// <param name="attachmentPaths">An array of the complete path to the attachment on the web server including 
        /// the file name. The path should be the local path instead of the website path, 
        /// i.e. c:\inetpub\uploads\info.doc
        /// </param>
        public static void SendEmailWithAttachments(string SMTPServer, string emailFrom, string emailFromName, string emailTo,
            string subject, string message, string[] attachmentPaths)
        {
            string strEmailFrom = "";		// The complete sender for the email (including the name if present)

            // Determine the Sender. Include the email and name if present.
            if (emailFromName != "")
                strEmailFrom = emailFromName + " <" + emailFrom + ">";
            else
                strEmailFrom = emailFrom;

            // Send the email
            try
            {
                SmtpClient smtpcMessage = new SmtpClient(SMTPServer);

                MailMessage mmMessage = new MailMessage(strEmailFrom, emailTo, subject, message);

                foreach (string path in attachmentPaths)
                {
                    if (path.Trim() != "")
                    {
                        if (System.IO.File.Exists(path))
                        {
                            Attachment aAttachment = new Attachment(path);
                            mmMessage.Attachments.Add(aAttachment);
                        }
                    }
                }

                smtpcMessage.Send(mmMessage);
            }
            catch
            {
                // Do nothing if there was an error as there is nothing we can do. 
                return;
            }

        }


	}
}

