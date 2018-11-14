using Microsoft.Exchange.WebServices.Data;
using SA.Utils;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using outlook = Microsoft.Office.Interop.Outlook;


namespace SA.Controllers
{
    public class OutlookController : Controller
    {
        // GET: Outlook
        public ActionResult Index()
        {
			string email = ConfigurationManager.AppSettings["UserEmail"];
			string password = ConfigurationManager.AppSettings["UserPassword"];

			EWSOutlook outlookService = new EWSOutlook(email, password);

			// Need to modify to allow accepting of folder type > Unit Test for methods later
			List<EmailMessage> inboxEmails = outlookService.InboxEmails();
			List<EmailMessage> folderEmails = outlookService.CustomerFolderEmails("Admin");

			// Using Timestamp to retrieve email within a period on a specific folder
			FolderId folderId = WellKnownFolderName.Inbox;

			DateTime startTimestamp = new DateTime(2018, 11, 12);
			TimeSpan startTime = new TimeSpan(18, 0, 0);
			startTimestamp = startTimestamp.Date + startTime;

			DateTime endTimestamp = new DateTime(2018, 11, 13);
			TimeSpan endTime = new TimeSpan(12, 0, 0);
			endTimestamp = endTimestamp.Date + endTime;
			
			List<EmailMessage> filterByDateInboxEmails = outlookService.StartEndTimestampInboxEmails(folderId, startTimestamp, endTimestamp);

			ViewBag.InboxEmails = inboxEmails;
			ViewBag.FolderEmails = folderEmails;
			ViewBag.FilteredEmails = filterByDateInboxEmails;

			return View();
        }

		public ActionResult Emails()
		{
			Outlook outlook = new Outlook();
			List<outlook.MailItem> inboxEmails = outlook.ExtractEmails(outlook.RetrieveInbox());

			ViewBag.InboxEmails = inboxEmails;

			return View();
		}
    }
}