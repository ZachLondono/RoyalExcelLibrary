using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Outlook = Microsoft.Office.Interop.Outlook;
using Excel = Microsoft.Office.Interop.Excel;

namespace RoyalExcelLibrary.ExportFormat {

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
			mailItem.To = args.To.Aggregate((a, b) => a += "; " + b);
			if (!(args.CC is null)) mailItem.CC = args.CC.Aggregate((a, b) => a += "; " + b);
			mailItem.Subject = args.Subject;
			mailItem.Body = args.Body;
			mailItem.SendUsingAccount = sendingAccount;

			if (!(args.Attachments is null)) {

				string tempfolder = System.IO.Path.GetTempPath();

				foreach (object attachmentSource in args.Attachments) {

					if (attachmentSource is string)
						mailItem.Attachments.Add(Source: attachmentSource);
					else if(attachmentSource is Excel.Worksheet) {
					
						string exportPath = $"{tempfolder}{(attachmentSource as Excel.Worksheet).Name}.pdf";
						(attachmentSource as Excel.Worksheet).ExportAsFixedFormat(Excel.XlFixedFormatType.xlTypePDF, Filename:exportPath);
						mailItem.Attachments.Add(Source: exportPath);
					
					} else if (attachmentSource is AttachmentArgs attArgs) {

                        string exportPath = $"{tempfolder}{attArgs.FileName}.pdf";
                        attArgs.Source.ExportAsFixedFormat(Excel.XlFixedFormatType.xlTypePDF, Filename: exportPath);
                        mailItem.Attachments.Add(Source: exportPath, DisplayName: attArgs.DisplayName);

                    }

                }

			}

			if (args.AutoSend)
				mailItem.Send();
			else mailItem.Display();

		}

	}

}