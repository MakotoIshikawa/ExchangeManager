using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExchangeManager;
using ExchangeManager.Extensions;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Ews = Microsoft.Exchange.WebServices.Data;
using ExtensionsLibrary.Extensions;

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
			var to = _username + ";" + "root@kariverification14.onmicrosoft.com";

			service.SendMail(to, subject, text);
		}

		[TestMethod]
		[Owner(nameof(ExchangeOnlineManager))]
		[TestCategory("送信")]
		public async Task 非同期でメールを送信する() {
			var service = new ExchangeOnlineManager(_username, _password);

			var subject = "Hello World!";
			var text = "これは、EWS Managed APIを使用して非同期で送信したメールです。";
			var to = _username + ";" + "root@kariverification14.onmicrosoft.com";

			await service.SendMailAsync(to, subject, text);
		}

		[TestMethod]
		[Owner(nameof(ExchangeOnlineManager))]
		[TestCategory("取得")]
		public void 予定を取得する() {
			var user = "root@kariverification14.onmicrosoft.com";
			var pass = "!QAZ2wsx";
			var service = new ExchangeOnlineManager(user, pass);

			// 開始時刻と終了時刻の値、および取得する予定の数を初期化します。
			var startDate = DateTime.Now;
			var endDate = startDate.AddDays(30);

			var appointments = service.FindAppointments(startDate, endDate);

			var sb = new StringBuilder();

			foreach (var a in appointments) {
				sb.AppendLine($"{a.Start:yyyy/MM/dd(ddd) HH:mm} ~ {a.End:yyyy/MM/dd(ddd) HH:mm} {a.Subject}");
			}

			var subject = $"[{user}] {startDate:yyyy/MM/dd(ddd)} ~ {endDate:yyyy/MM/dd(ddd)} の予定";
			var text = sb.ToString();
			var to = $"{_username};{user};";

			service.SendMail(to, subject, text);
		}

		[TestMethod]
		[Owner(nameof(ExchangeOnlineManager))]
		[TestCategory("取得")]
		public async Task 非同期で予定を取得する() {
			var user = "root@kariverification14.onmicrosoft.com";
			var pass = "!QAZ2wsx";
			var service = new ExchangeOnlineManager(user, pass);

			// 開始時刻と終了時刻の値、および取得する予定の数を初期化します。
			var startDate = DateTime.Now;
			var endDate = startDate.AddDays(30);

			var appointments = await service.FindAppointmentsAsync(startDate, endDate);

			var sb = new StringBuilder();

			foreach (var a in appointments) {
				sb.AppendLine($"{a.Start:yyyy/MM/dd(ddd) HH:mm} ~ {a.End:yyyy/MM/dd(ddd) HH:mm} {a.Subject}");
			}

			var subject = $"[{user}] {startDate:yyyy/MM/dd(ddd)} ~ {endDate:yyyy/MM/dd(ddd)} の予定";
			var text = sb.ToString();
			var to = $"{_username};{user};";

			await service.SendMailAsync(to, subject, text);
		}

		[TestMethod]
		[Owner(nameof(ExchangeOnlineManager))]
		[TestCategory("取得")]
		public void 空き時間を取得する() {
			// 出席者のコレクションを作成します。 
			var attendees = new List<Ews.AttendeeInfo> {
				{ "root@kariverification14.onmicrosoft.com", Ews.MeetingAttendeeType.Organizer },		// 主催者
				{ "ishikawm@kariverification14.onmicrosoft.com", Ews.MeetingAttendeeType.Required },	// 必須
				//{ "karikomi@kariverification14.onmicrosoft.com", Ews.MeetingAttendeeType.Optional },	// 任意
				{ "chiakimi@kariverification14.onmicrosoft.com", Ews.MeetingAttendeeType.Optional },	// 任意
				{ "conference_f29_01@kariverification14.onmicrosoft.com", Ews.MeetingAttendeeType.Room },	// 会議室
				{ "conference_f29_02@kariverification14.onmicrosoft.com", Ews.MeetingAttendeeType.Room },	// 会議室
			};

			var date = new DateTime(2017, 04, 17);
			var startTime = date;
			var endTime = date.AddDays(4) - new TimeSpan(1);

			var service = new ExchangeOnlineManager(_username, _password);
			var sc = new ExchangeScheduler(service, startTime, endTime, attendees);

			var meetingDuration = 60;

			var suggestions = sc.GetSuggestions();

			var sb = new StringBuilder();

			// 提案された会議時間を表示します。
			var str = attendees.Select(a => $"\t{a.SmtpAddress}").Join("\n");
			sb.AppendLine($"アドレス :\n{str}").AppendLine();

			sb.AppendLine("--------------------------------------------------------------------------------");

			foreach (var suggestion in suggestions) {
				sb.AppendLine($"提案日: {suggestion.Date:d}");
				sb.AppendLine($"推奨される会議時間:");
				foreach (var ts in suggestion.TimeSuggestions) {
					var tim = ts.MeetingTime;
					sb.AppendLine($"\t{tim:t} ~ {tim.AddMinutes(meetingDuration):t}");
				}

				sb.AppendLine();
			}

			sb.AppendLine();

			var infos = sc.GetUserAvailabilities();

			infos.ForEach(info => {
				sb.AppendLine($"[{info.Key}]:");

				// 出席者のカレンダーイベントのコレクションを取得します。
				foreach (var ev in info.Value) {
					sb.AppendLine($"\t{ev.StartTime:yyyy/MM/dd(ddd) HH:mm} ~ {ev.EndTime:yyyy/MM/dd(ddd) HH:mm} [{ev.FreeBusyStatus}] : {ev.Details?.GetPropertiesString()}");
				}

				sb.AppendLine();
			});

			sb.AppendLine("--------------------------------------------------------------------------------");

			var user = "root@kariverification14.onmicrosoft.com";
			var subject = $"{startTime:yyyy/MM/dd(ddd)} ~ {endTime:yyyy/MM/dd(ddd)} の推奨される会議時間";
			var text = sb.ToString();
			var to = $"{_username};{user};";

			service.SendMail(to, subject, text);
		}


		[TestMethod]
		[Owner(nameof(ExchangeOnlineManager))]
		[TestCategory("取得")]
		public async Task 非同期で空き時間を取得する() {
			var service = new ExchangeOnlineManager(_username, _password);

			var meetingDuration = 60;

			// 出席者のコレクションを作成します。 
			var attendees = new List<Ews.AttendeeInfo> {
				{ "root@kariverification14.onmicrosoft.com", Ews.MeetingAttendeeType.Organizer },		// 主催者
				{ "ishikawm@kariverification14.onmicrosoft.com", Ews.MeetingAttendeeType.Required },	// 必須
				{ "karikomi@kariverification14.onmicrosoft.com", Ews.MeetingAttendeeType.Optional },	// 任意
				{ "chiakimi@kariverification14.onmicrosoft.com", Ews.MeetingAttendeeType.Optional },	// 任意
				{ "conference_f29_01@kariverification14.onmicrosoft.com", Ews.MeetingAttendeeType.Room },	// 会議室
				{ "conference_f29_02@kariverification14.onmicrosoft.com", Ews.MeetingAttendeeType.Room },	// 会議室
			};

			var date = new DateTime(2017, 04, 17);
			var startTime = date;
			var endTime = date.AddDays(4);
			var lastTime = endTime - new TimeSpan(1);

			var results = await service.GetUserAvailabilityAsync(attendees, startTime, endTime, meetingDuration: meetingDuration);

			var sb = new StringBuilder();

			// 提案された会議時間を表示します。
			var str = attendees.Select(a => $"\t{a.SmtpAddress}").Join("\n");
			sb.AppendLine($"アドレス :\n{str}").AppendLine();

			sb.AppendLine("--------------------------------------------------------------------------------");

			foreach (var suggestion in results.Suggestions) {
				sb.AppendLine($"提案日: {suggestion.Date:yyyy/MM/dd(ddd)}");
				sb.AppendLine($"推奨される会議時間:");
				foreach (var ts in suggestion.TimeSuggestions) {
					var tim = ts.MeetingTime;
					sb.AppendLine($"\t{tim:t} ~ {tim.AddMinutes(meetingDuration):t}");
				}

				sb.AppendLine();
			}

			var infos = attendees.Zip(results.AttendeesAvailability, (at, av) => new {
				Attendee = at,
				Availability = av,
			});

			infos.ForEach(info => {
				sb.AppendLine($"[{info.Attendee.SmtpAddress}]:");

				// 出席者のカレンダーイベントのコレクションを取得します。
				foreach (var ev in info.Availability.CalendarEvents) {
					sb.AppendLine($"\t{ev.StartTime:yyyy/MM/dd(ddd) HH:mm} ~ {ev.EndTime:yyyy/MM/dd(ddd) HH:mm} [{ev.FreeBusyStatus}] : {ev.Details?.GetPropertiesString()}");
				}

				sb.AppendLine();

				if (info.Availability.MergedFreeBusyStatus?.Any() ?? false) {
					// 出席者の空き/会議中状態を結合したコレクションを取得します。
					foreach (var fb in info.Availability.MergedFreeBusyStatus) {
						sb.AppendLine($"\tStatus {fb}");
					}

					sb.AppendLine();
				}
			});

			sb.AppendLine("--------------------------------------------------------------------------------");

			var user = "root@kariverification14.onmicrosoft.com";
			var subject = $"{startTime:yyyy/MM/dd(ddd)} ~ {lastTime:yyyy/MM/dd(ddd)} の推奨される会議時間 (非同期処理で取得しました)";
			var text = sb.ToString();
			var to = $"{_username};{user};";

			await service.SendMailAsync(to, subject, text);
		}

		#endregion
	}
}
