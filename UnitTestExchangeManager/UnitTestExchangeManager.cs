using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExchangeManager;
using ExchangeManager.Extensions;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Ews = Microsoft.Exchange.WebServices.Data;
using ExtensionsLibrary.Extensions;
using UnitTestExchangeManager.Properties;

namespace UnitTestExchangeManager {
	[TestClass]
	public class UnitTestExchangeManager {
		#region フィールド

		private static string _username = Settings.Default.UserName;
		private static string _password = Settings.Default.Password;

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
			var service = new ExchangeOnlineManager(_username, _password);

			// 開始時刻と終了時刻の値、および取得する予定の数を初期化します。
			var startDate = DateTime.Now;
			var endDate = startDate.AddDays(30);

			var appointments = service.FindAppointments(startDate, endDate);

			var sb = new StringBuilder();

			foreach (var a in appointments) {
				sb.AppendLine($"{a.Start:yyyy/MM/dd(ddd) HH:mm} ~ {a.End:yyyy/MM/dd(ddd) HH:mm} {a.Subject}");
			}

			var user = "root@kariverification14.onmicrosoft.com";
			var subject = $"[{_username}] {startDate:yyyy/MM/dd(ddd)} ~ {endDate:yyyy/MM/dd(ddd)} の予定";
			var text = sb.ToString();
			var to = $"{_username};{user};";

			service.SendMail(to, subject, text);
		}

		[TestMethod]
		[Owner(nameof(ExchangeOnlineManager))]
		[TestCategory("取得")]
		public async Task 非同期で予定を取得する() {
			var service = new ExchangeOnlineManager(_username, _password);

			// 開始時刻と終了時刻の値、および取得する予定の数を初期化します。
			var startDate = DateTime.Now;
			var endDate = startDate.AddDays(30);

			var appointments = await service.FindAppointmentsAsync(startDate, endDate);

			var sb = new StringBuilder();

			foreach (var a in appointments) {
				sb.AppendLine($"{a.Start:yyyy/MM/dd(ddd) HH:mm} ~ {a.End:yyyy/MM/dd(ddd) HH:mm} {a.Subject}");
			}

			var user = "root@kariverification14.onmicrosoft.com";
			var subject = $"[{_username}] {startDate:yyyy/MM/dd(ddd)} ~ {endDate:yyyy/MM/dd(ddd)} の予定";
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

			var date = new DateTime(2017, 04, 24);
			var start = date;
			var end = date.AddDays(3);

			var service = new ExchangeOnlineManager(_username, _password);
			var sc = new ExchangeScheduler(service, start, end, attendees) {
				GoodSuggestionThreshold = 49,
				MaximumNonWorkHoursSuggestionsPerDay = 0,
				MaximumSuggestionsPerDay = 2,
				MeetingDuration = 60,
				MinimumSuggestionQuality = Ews.SuggestionQuality.Good,
				RequestedFreeBusyView = Ews.FreeBusyViewType.FreeBusy,
			};

			var suggestions = sc.GetSuggestions();

			var sb = new StringBuilder();

			// 提案された会議時間を表示します。
			var str = attendees.Select(a => $"\t{a.SmtpAddress}").Join("\n");
			sb.AppendLine($"アドレス :\n{str}").AppendLine();

			sb.AppendLine("--------------------------------------------------------------------------------");

			foreach (var s in suggestions) {
				sb.AppendLine($"提案日: {s.Key:d}");
				sb.AppendLine($"推奨される会議時間:");
				foreach (var t in s.Value) {
					sb.AppendLine($"\t{t.StartTime:t} ~ {t.EndTime:t}");
				}

				sb.AppendLine();
			}

			sb.AppendLine();

			var availabilities = sc.GetUserAvailabilities();

			availabilities.ForEach(info => {
				sb.AppendLine($"[{info.Key}]:");

				// 出席者のカレンダーイベントのコレクションを取得します。
				foreach (var ev in info.Value) {
					sb.AppendLine($"\t{ev.StartTime:yyyy/MM/dd(ddd) HH:mm} ~ {ev.EndTime:yyyy/MM/dd(ddd) HH:mm} [{ev.FreeBusyStatus}] : {ev.Details?.GetPropertiesString()}");
				}

				sb.AppendLine();
			});

			sb.AppendLine("--------------------------------------------------------------------------------");

			var user = "root@kariverification14.onmicrosoft.com";
			var subject = $"{sc.StartTime:yyyy/MM/dd(ddd)} ~ {sc.EndTime:yyyy/MM/dd(ddd)} の推奨される会議時間";
			var text = sb.ToString();
			var to = $"{_username};{user};";

			service.SendMail(to, subject, text);
		}

		[TestMethod]
		[Owner(nameof(ExchangeOnlineManager))]
		[TestCategory("取得")]
		public async Task 非同期で空き時間を取得する() {
			// 出席者のコレクションを作成します。 
			var attendees = new List<Ews.AttendeeInfo> {
				{ "root@kariverification14.onmicrosoft.com", Ews.MeetingAttendeeType.Organizer },		// 主催者
				{ "ishikawm@kariverification14.onmicrosoft.com", Ews.MeetingAttendeeType.Required },	// 必須
				//{ "karikomi@kariverification14.onmicrosoft.com", Ews.MeetingAttendeeType.Optional },	// 任意
				{ "chiakimi@kariverification14.onmicrosoft.com", Ews.MeetingAttendeeType.Optional },	// 任意
				{ "conference_f29_01@kariverification14.onmicrosoft.com", Ews.MeetingAttendeeType.Room },	// 会議室
				{ "conference_f29_02@kariverification14.onmicrosoft.com", Ews.MeetingAttendeeType.Room },	// 会議室
			};

			var date = new DateTime(2017, 04, 24);
			var start = date;
			var end = date.AddDays(3);

			var service = new ExchangeOnlineManager(_username, _password);
			var sc = new ExchangeScheduler(service, start, end, attendees) {
				GoodSuggestionThreshold = 49,
				MaximumNonWorkHoursSuggestionsPerDay = 0,
				MaximumSuggestionsPerDay = 2,
				MeetingDuration = 60,
				MinimumSuggestionQuality = Ews.SuggestionQuality.Good,
				RequestedFreeBusyView = Ews.FreeBusyViewType.FreeBusy,
			};

			var suggestions = await sc.GetSuggestionsAsync();

			var sb = new StringBuilder();

			// 提案された会議時間を表示します。
			var str = attendees.Select(a => $"\t{a.SmtpAddress}").Join("\n");
			sb.AppendLine($"アドレス :\n{str}").AppendLine();

			sb.AppendLine("--------------------------------------------------------------------------------");

			foreach (var s in suggestions) {
				sb.AppendLine($"提案日: {s.Key:d}");
				sb.AppendLine($"推奨される会議時間:");
				foreach (var t in s.Value) {
					sb.AppendLine($"\t{t.StartTime:t} ~ {t.EndTime:t}");
				}

				sb.AppendLine();
			}

			sb.AppendLine();

			var availabilities = await sc.GetUserAvailabilitiesAsync();

			availabilities.ForEach(info => {
				sb.AppendLine($"[{info.Key}]:");

				// 出席者のカレンダーイベントのコレクションを取得します。
				foreach (var ev in info.Value) {
					sb.AppendLine($"\t{ev.StartTime:yyyy/MM/dd(ddd) HH:mm} ~ {ev.EndTime:yyyy/MM/dd(ddd) HH:mm} [{ev.FreeBusyStatus}] : {ev.Details?.GetPropertiesString()}");
				}

				sb.AppendLine();
			});

			sb.AppendLine("--------------------------------------------------------------------------------");

			var user = "root@kariverification14.onmicrosoft.com";
			var subject = $"{sc.StartTime:yyyy/MM/dd(ddd)} ~ {sc.EndTime:yyyy/MM/dd(ddd)} の推奨される会議時間";
			var text = sb.ToString();
			var to = $"{_username};{user};";

			await service.SendMailAsync(to, subject, text);
		}

		#endregion
	}
}
