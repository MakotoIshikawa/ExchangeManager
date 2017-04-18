using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExchangeManager.Extensions;
using ExchangeManager.Interface;
using Microsoft.Exchange.WebServices.Data;

namespace ExchangeManager {
	public class Program {
		private static string _username = @"ishikawm@kariverification14.onmicrosoft.com";
		private static string _password = @"Ishikawam!";

		public static void Main(string[] args) {
			var service = new ExchangeOnlineManager(_username, _password);

			GetSuggestedMeetingTimesAndFreeBusyInfo(service);
		}

		private static void GetSuggestedMeetingTimesAndFreeBusyInfo(IExchangeManager service) {
			var meetingDuration = 60;

			// 出席者のコレクションを作成します。 
			var attendees = new List<AttendeeInfo> {
				{ "mack@contoso.com", MeetingAttendeeType.Organizer },	// 主催者
				{ "sadie@contoso.com", MeetingAttendeeType.Required },	// 必須
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
	}

	public static partial class Extension {
	}
}
