using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Reflection;
using outlook = Microsoft.Office.Interop.Outlook;

namespace SA.Utils
{
	public class Outlook
	{
		outlook.Application oApp;
		outlook.NameSpace oNS;

		public Outlook()
		{
			if (!OutlookStatus())
			{
				oApp = new outlook.Application();
				oNS = oApp.GetNamespace("mapi");
				oNS.Logon(Missing.Value, Missing.Value, false, true);
			}
			else
			{
				oApp = new outlook.Application();
				oNS = oApp.GetNamespace("mapi");
				oNS.Logon(Missing.Value, Missing.Value, false, false);
			}
		}

		public outlook.MAPIFolder RetrieveInbox()
		{
			outlook.MAPIFolder oInbox = oNS.GetDefaultFolder(outlook.OlDefaultFolders.olFolderInbox);

			return oInbox;
		}

		public outlook.MAPIFolder RetrieveFolder(string folderName)
		{
			outlook.MAPIFolder oFolder = null;

			outlook.MAPIFolder oInbox = oNS.GetDefaultFolder(outlook.OlDefaultFolders.olFolderInbox);

			try
			{
				oFolder= oInbox.Folders[folderName];
				outlook.Items oItems = oFolder.Items;		
			}
			catch
			{
				// Error handling 
			}
			return oFolder;
		}


		public List<outlook.MailItem> ExtractEmails(outlook.MAPIFolder oFolder)
		{
			List<outlook.MailItem> oMails = null;

			outlook.Items oItems = oFolder.Items;

			foreach (object item in oItems)
			{
				if (item is outlook.MailItem)
				{
					oMails.Add((outlook.MailItem)item);
				}
			}
			return oMails;
		}

		private bool OutlookStatus()
		{
			if (System.Diagnostics.Process.GetProcessesByName("OUTLOOK").Any()) { return true; }
			else { return false; }
		}
	}
}