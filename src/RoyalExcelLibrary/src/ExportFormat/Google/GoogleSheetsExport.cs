using RoyalExcelLibrary.ExcelUI.Models;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;

namespace RoyalExcelLibrary.ExcelUI.ExportFormat.Google {
	public abstract class GoogleSheetsExport : IGoogleSheetsExport {

		protected List<object> Data;

		protected GoogleSheetsExport() {
			Data = new List<object>();
		}

		public abstract void ExportOrder(Order order);

		protected void ExportCurrentData() {

			ProcessStartInfo startInfo = new ProcessStartInfo();
			startInfo.CreateNoWindow = true;
			startInfo.UseShellExecute = false;
			startInfo.WindowStyle = ProcessWindowStyle.Hidden;
			startInfo.FileName = "R:\\DB ORDERS\\GoogleSheetsExe\\publish\\GoogleSheetsUpdater.exe";
			string argStr = "";
			foreach (object arg in Data) {
				argStr += $"\"{arg.ToString()}\" ";
			}
			Debug.WriteLine($"Running google sheet updater: '{startInfo.FileName} {argStr}'");
			startInfo.Arguments = argStr;

			try {
				using (Process process = Process.Start(startInfo)) {

					string outputText = "";

					process.OutputDataReceived += new DataReceivedEventHandler ( 
						delegate (object sender, DataReceivedEventArgs args) {
							using (StreamReader output = process.StandardOutput) {
								outputText = output.ReadToEnd();
							}
						}
					);

					Debug.WriteLine($"Google sheet updater stdout: '{outputText}'");

					process.WaitForExit();
				}
			} catch (Exception e) {
				throw new Exception ("Error while tracking order on google sheet", e);
				throw;
			}

		}

	}

}
