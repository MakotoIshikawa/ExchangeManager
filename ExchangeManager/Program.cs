using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Exchange.WebServices.Data;

namespace ExchangeManager {
	class Program {
		static void Main(string[] args) {
			var username = @"ishikawm@kariverification14.onmicrosoft.com";
			var password = @"Ishikawam!";

			var service = new ExchangeService(ExchangeVersion.Exchange2013_SP1) {
				Credentials = new WebCredentials(username, password),
				UseDefaultCredentials = false,
				TraceEnabled = true,
				TraceFlags = TraceFlags.All,
			};

			service.AutodiscoverUrl(username, url => {
				// 検証コールバックのデフォルトは、URLを拒否することです。
				var result = false;

				var redirectionUri = new Uri(url);

				// リダイレクトURLの内容を検証します。
				// この単純な検証コールバックでは、HTTPSを使用して認証資格情報を暗号化する場合、
				// リダイレクトURLは有効と見なされます。
				if (redirectionUri.Scheme == "https") {
					result = true;
				}

				return result;
			});

			var subject = "HelloWorld";
			var text = "これは、EWS Managed APIを使用して送信した最初のメールです。";
			var isRichText = false;

			var email = new EmailMessage(service) {
				Subject = subject,
				Body = new MessageBody(isRichText ? BodyType.HTML : BodyType.Text, text),
			};
			email.ToRecipients.Add(username);

			email.Send();

			// Initialize values for the start and end times, and the number of appointments to retrieve.
			var startDate = DateTime.Now;
			var endDate = startDate.AddDays(30);
			int numAppts = 5;

			// Initialize the calendar folder object with only the folder ID. 
			var calendar = CalendarFolder.Bind(service, WellKnownFolderName.Calendar, new PropertySet());

			// Set the start and end time and number of appointments to retrieve.
			var cView = new CalendarView(startDate, endDate, numAppts);

			// Limit the properties returned to the appointment's subject, start time, and end time.
			cView.PropertySet = new PropertySet(AppointmentSchema.Subject, AppointmentSchema.Start, AppointmentSchema.End);

			// Retrieve a collection of appointments by using the calendar view.
			FindItemsResults<Appointment> appointments = calendar.FindAppointments(cView);

			Console.WriteLine("\nThe first " + numAppts + " appointments on your calendar from " + startDate.Date.ToShortDateString() +
							  " to " + endDate.Date.ToShortDateString() + " are: \n");

			foreach (Appointment a in appointments) {
				Console.Write("Subject: " + a.Subject.ToString() + " ");
				Console.Write("Start: " + a.Start.ToString() + " ");
				Console.Write("End: " + a.End.ToString());
				Console.WriteLine();
			}
		}
	}
}
