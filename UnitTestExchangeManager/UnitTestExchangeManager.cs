using System;
using System.Collections.Generic;
using System.Diagnostics;
using ExchangeManager;
using ExchangeManager.Extensions;
using Microsoft.Exchange.WebServices.Data;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace UnitTestExchangeManager {
	[TestClass]
	public class UnitTestExchangeManager {
		#region フィールド

		private static string _username = @"ishikawm@kariverification14.onmicrosoft.com";
		private static string _password = @"Ishikawam!";

		#endregion

		#region メソッド

		[TestMethod]
		[Owner(nameof(ExchangeOnlineManager))]
		[TestCategory("送信")]
		public void メールを送信する() {
			var service = new ExchangeOnlineManager(_username, _password);

			var subject = "Hello World!";
			var text = "これは、EWS Managed APIを使用して送信した最初のメールです。";
			var to = _username + ";" + "ishikawm@fsi.co.jp";

			service.SendMail(to, subject, text);
		}

		[TestMethod]
		[Owner(nameof(ExchangeOnlineManager))]
		[TestCategory("取得")]
		public void 予定を取得する() {
			var service = new ExchangeOnlineManager(_username, _password);

			// 開始時刻と終了時刻の値、および取得する予定の数を初期化します。
			var startDate = DateTime.Now;
			var endDate = startDate.AddDays(30);
			var maxItemsReturned = 5;

			var appointments = service.FindAppointments(startDate, endDate, maxItemsReturned);

			Debug.WriteLine("\nThe first " + maxItemsReturned + " appointments on your calendar from " + startDate.Date.ToShortDateString() +
							  " to " + endDate.Date.ToShortDateString() + " are: \n");

			foreach (var a in appointments) {
				Debug.Write("Subject: " + a.Subject.ToString() + " ");
				Debug.Write("Start: " + a.Start.ToString() + " ");
				Debug.WriteLine("End: " + a.End.ToString());
			}

			Assert.IsTrue(true);
		}

		[TestMethod]
		[Owner(nameof(ExchangeOnlineManager))]
		[TestCategory("取得")]
		public void 空き時間を取得する() {
			var service = new ExchangeOnlineManager(_username, _password);

			var meetingDuration = 60;

			// 出席者のコレクションを作成します。 
			var attendees = new List<AttendeeInfo> {
				{ "root@kariverification14.onmicrosoft.com", MeetingAttendeeType.Organizer },		// 主催者
				{ "ishikawm@kariverification14.onmicrosoft.com", MeetingAttendeeType.Required },	// 必須
				{ "conference_f29_01@kariverification14.onmicrosoft.com", MeetingAttendeeType.Room },	// 会議室
			};

			var now = DateTime.Now;
			var startTime = now.AddDays(1);
			var endTime = now.AddDays(2);

			var goodSuggestionThreshold = 49;
			var maximumNonWorkHoursSuggestionsPerDay = 0;
			var maximumSuggestionsPerDay = 2;

			var results = service.GetUserAvailability(attendees, startTime, endTime, goodSuggestionThreshold, maximumNonWorkHoursSuggestionsPerDay, maximumSuggestionsPerDay, meetingDuration);

			// 提案された会議時間を表示します。
			Debug.WriteLine($"Availability for {attendees[0].SmtpAddress} and {attendees[1].SmtpAddress}");

			foreach (var suggestion in results.Suggestions) {
				Debug.WriteLine($"Suggested date: {suggestion.Date.ToShortDateString()}\n");
				Debug.WriteLine($"Suggested meeting times:\n");
				foreach (var ts in suggestion.TimeSuggestions) {
					var tim = ts.MeetingTime;
					Debug.WriteLine($"\t{tim.ToShortTimeString()} - {tim.AddMinutes(meetingDuration).ToShortTimeString()}\n");
				}

				int i = 0;

				// 空き時間を表示します。
				foreach (var availability in results.AttendeesAvailability) {
					Debug.WriteLine($"Availability information for {attendees[i].SmtpAddress}:\n");

					foreach (var calEvent in availability.CalendarEvents) {
						Debug.WriteLine($"\tBusy from {calEvent.StartTime} to {calEvent.EndTime} \n");
					}

					i++;
				}
			}
		}

		#endregion
	}
}
