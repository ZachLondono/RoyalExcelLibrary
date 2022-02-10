using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Outlook = Microsoft.Office.Interop.Outlook;
using Excel = Microsoft.Office.Interop.Excel;

namespace RoyalExcelLibrary.ExcelUI.ExportFormat {

	public struct EmailArgs {
		public string From { get; set; }
		public string[] To { get; set; }
		public string[] CC { get; set; }
		public string Subject { get; set; }
		public string Body { get; set; }
		public object[] Attachments { get; set; }
		public bool AutoSend { get; set; }
	}

	public struct AttachmentArgs {
		public Excel.Worksheet Source { get; set; }
		public string FileName { get; set; }
		public string DisplayName { get; set; }

	}

	public class OutlookEmailExport {

		public static void SendEmail(EmailArgs args) {

			if (args.From is null) throw new InvalidOperationException("No email sender specified");
			if (args.To is null) throw new InvalidOperationException("No email recipient specified");

			Outlook.Application olkApp = new Outlook.Application();
			Outlook.Accounts accounts = olkApp.Session.Accounts;
			Outlook.Account sendingAccount = null;
			foreach(Outlook.Account account in accounts) {
				if (account.SmtpAddress.Equals(args.From)) {
					sendingAccount = account;
					break;
				}
			}

			if (sendingAccount is null)
				throw new InvalidOperationException($"Unable to access email '{args.From}'\nMake sure you are logged in to this email in outlook");

			Outlook.MailItem mailItem = olkApp.CreateItem(Outlook.OlItemType.olMailItem);
			if (!(args.To is null) && args.To.Length != 0) mailItem.To = args.To.Aggregate((a, b) => a += "; " + b); // Converts the 'To' List if emails to a single string with all emails seperated by ';'
			if (!(args.CC is null) && args.CC.Length != 0) mailItem.CC = args.CC.Aggregate((a, b) => a += "; " + b); // Converts the 'CC' List if emails to a single string with all emails seperated by ';'
			mailItem.Subject = args.Subject;

			// Need to display the email in order to generate the signature
			mailItem.Display();
			mailItem.HTMLBody = InsertIntoExistingBody(mailItem.HTMLBody, args.Body);
			mailItem.SendUsingAccount = sendingAccount;

			if (!(args.Attachments is null)) {

				string tempfolder = System.IO.Path.GetTempPath();

				foreach (object attachmentSource in args.Attachments) {

					if (attachmentSource is string) // If the attachment source is a string, it is a file path
						mailItem.Attachments.Add(Source: attachmentSource);
					else if(attachmentSource is Excel.Worksheet) {
					 
						// If the attachment source is an excel sheet, print the sheet to a pdf in the temp directory and email that file

						string exportPath = $"{tempfolder}\\{(attachmentSource as Excel.Worksheet).Name}.pdf";
						(attachmentSource as Excel.Worksheet).ExportAsFixedFormat(Excel.XlFixedFormatType.xlTypePDF, Filename:exportPath);
						mailItem.Attachments.Add(Source: exportPath);
					
					} else if (attachmentSource is AttachmentArgs attArgs) {

                        string exportPath = $"{tempfolder}\\{attArgs.FileName}.pdf";
                        attArgs.Source.ExportAsFixedFormat(Excel.XlFixedFormatType.xlTypePDF, Filename: exportPath);
                        mailItem.Attachments.Add(Source: exportPath, DisplayName: attArgs.DisplayName);

                    }

                }

			}

			if (args.AutoSend)
				mailItem.Send();

		}

		/// <summary>
		/// Inserts new html into the top of the body of the existing html. When sending an email with outlook, the default signature is generated upon initial email creation, but then overwritten when setting the html body. In order to maintain the signiture, the body of the email must be inserted into the body tag of the existing html
		/// </summary>
		/// <param name="existingBody">Existing html with a body tag</param>
		/// <param name="newComponent">New html component, will be surrounded by a span tag in the resulting html</param>
		/// <returns></returns>
		public static string InsertIntoExistingBody(string existingBody, string newComponent) {

			int btagStartIndex = existingBody
									.ToLower()
									// Don't include the closing bracket '>' because the body tag may have additional attributes
									.IndexOf("<body");

			int btagEndIndex = btagStartIndex + existingBody
												.Substring(btagStartIndex)
												.IndexOf(">");

			// The prefix is all the html code up until the end of the <body> tag, which may include a number of attributes
			string prefix = existingBody.Substring(0, btagEndIndex + 1);

			// The suffix is all the html code after the end of the <body> tag
			string suffix = existingBody.Substring(btagEndIndex + 1);

			StringBuilder builder = new StringBuilder(existingBody.Length + newComponent.Length);
			builder.Append(prefix);
			builder.Append("<span>");
			builder.Append(newComponent);
			builder.Append("</span>");
			builder.Append(suffix);
			
			return builder.ToString();
        }

	}

}