using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Net;
using System.Web;
using Microsoft.Exchange.WebServices.Data;
using SA.Utils;

namespace SA.Utils
{
	public class EWSOutlook
	{
		private ExchangeService service;

		public EWSOutlook(string userEmail, string userPassword)
		{
			ServicePointManager.ServerCertificateValidationCallback = EWSOutlookCertificateValidation.CertificateValidationCallBack;
			ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2013)
			{
				Credentials = new WebCredentials(userEmail, userPassword),

				TraceEnabled = true,
				TraceFlags = TraceFlags.All,
				EnableScpLookup = true
			};

			service.AutodiscoverUrl(userEmail, EWSOutlookCertificateValidation.RedirectionUrlValidationCallback);

			this.service = service;
		}

		public List<EmailMessage> InboxEmails()
		{
			int offset = 0;
			int pageSize = 50;
			ItemView view = new ItemView(pageSize, offset, OffsetBasePoint.Beginning)
			{
				PropertySet = PropertySet.FirstClassProperties //A PropertySet with the explicit properties you want goes here
			};

			FindItemsResults<Item> findResults = service.FindItems(WellKnownFolderName.Inbox, view);
			List<EmailMessage> emails = new List<EmailMessage>();

			foreach (var item in findResults.Items)
			{
				emails.Add((EmailMessage)item);
			}
			return emails;
		}

		public List<EmailMessage> StartEndTimestampInboxEmails(FolderId folderId, DateTime startTimestamp, DateTime endTimestamp)
		{
			int offset = 0;
			int pageSize = 50;
			ItemView view = new ItemView(pageSize, offset, OffsetBasePoint.Beginning)
			{
				PropertySet = PropertySet.FirstClassProperties //A PropertySet with the explicit properties you want goes here
			};

			List<EmailMessage> result = StartEndTimeEmailFilter(view, folderId, startTimestamp, endTimestamp);

			return result;
		}

		public List<EmailMessage> CustomerFolderEmails(string folderName)
		{
			int offset = 0;
			int pageSize = 50;
			ItemView view = new ItemView(pageSize, offset, OffsetBasePoint.Beginning)
			{
				PropertySet = PropertySet.FirstClassProperties
			};

			FolderId folderId = FolderSearch(folderName)[0].Id;

			FindItemsResults<Item> findResults = service.FindItems(folderId, view);

			return convertResultsToEmail(findResults);
		}

		private List<Folder> FolderSearch(string folderName)
		{
			FolderView view = new FolderView(100);
			view.PropertySet = new PropertySet(BasePropertySet.IdOnly);
			view.PropertySet.Add(FolderSchema.DisplayName);
			view.Traversal = FolderTraversal.Deep;

			SearchFilter.ContainsSubstring folderNameFilter = new SearchFilter.ContainsSubstring(
				FolderSchema.DisplayName,
				folderName,
				ContainmentMode.Substring,
				ComparisonMode.IgnoreCase
				);

			FindFoldersResults results = service.FindFolders(WellKnownFolderName.Root, folderNameFilter, view);

			List<Folder> folders = new List<Folder>();

			foreach (Folder item in results)
			{
				folders.Add(item);
			}
			return folders;
		}

		private List<EmailMessage> StartEndTimeEmailFilter(ItemView view, FolderId folderId, DateTime startTimeStamp, DateTime endTimeStamp)
		{
			List<Item> results = new List<Item>();

			SearchFilter.IsLessThanOrEqualTo earlierThan = new SearchFilter.IsLessThanOrEqualTo(ItemSchema.DateTimeReceived, endTimeStamp);
			SearchFilter.IsGreaterThanOrEqualTo laterThan = new SearchFilter.IsGreaterThanOrEqualTo(ItemSchema.DateTimeReceived, startTimeStamp);

			FindItemsResults<Item> earlierThanResults = service.FindItems(folderId, earlierThan, view);
			FindItemsResults<Item> laterThanResults = service.FindItems(folderId, laterThan, view);

			foreach (Item laterItem in laterThanResults)
			{
				foreach (Item earlierItem in earlierThanResults)
				{
					if (Convert.ToString(laterItem.Id).Equals(Convert.ToString(earlierItem.Id)))
					{
						results.Add(earlierItem);
						break;
					}
				}
			}
			return convertItemToEmail(results);
		}

		private List<EmailMessage> convertItemToEmail(List<Item> items)
		{
			List<EmailMessage> emails = new List<EmailMessage>();

			foreach (EmailMessage item in items)
			{
				emails.Add(item);
			}
			return emails;
		}

		private List<EmailMessage> convertResultsToEmail(FindItemsResults<Item> items)
		{
			List<EmailMessage> emails = new List<EmailMessage>();

			foreach (EmailMessage item in items.Items)
			{
				emails.Add(item);
			}
			return emails;
		}

	}
}