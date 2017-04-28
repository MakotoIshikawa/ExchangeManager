using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExchangeManager;
using ExchangeManager.Extensions;
using ExtensionsLibrary.Extensions;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using UnitTestExchangeManager.Properties;
using Ews = Microsoft.Exchange.WebServices.Data;

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
				{ "chiakimi@kariverification14.onmicrosoft.com", Ews.MeetingAttendeeType.Optional },	// 任意
				{ "conference_f29_01@kariverification14.onmicrosoft.com", Ews.MeetingAttendeeType.Room },	// 会議室
				{ "conference_f29_02@kariverification14.onmicrosoft.com", Ews.MeetingAttendeeType.Room },	// 会議室
			};

			var start = new DateTime(2017, 04, 24);
			var end = new DateTime(2017, 04, 28);

			var service = new ExchangeOnlineManager(_username, _password);
			var sc = new ExchangeScheduler(service, start, end, attendees) {
				GoodSuggestionThreshold = 49,
				MaximumNonWorkHoursSuggestionsPerDay = 8,
				MaximumSuggestionsPerDay = 8,
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
			var subject = $"{sc.StartTime:yyyy/MM/dd(ddd)} ~ {sc.LastTime:yyyy/MM/dd(ddd)} の推奨される会議時間";
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
				{ "chiakimi@kariverification14.onmicrosoft.com", Ews.MeetingAttendeeType.Optional },	// 任意
			};

			var start = DateTime.Today;
			var end = start.AddDays(7);

			var service = new ExchangeOnlineManager(_username, _password);
			var sc = new ExchangeScheduler(service, start, end, attendees) {
				GoodSuggestionThreshold = 49,
				MaximumNonWorkHoursSuggestionsPerDay = 8,
				MaximumSuggestionsPerDay = 8,
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
			var subject = $"{sc.StartTime:yyyy/MM/dd(ddd)} ~ {sc.LastTime:yyyy/MM/dd(ddd)} の推奨される会議時間";
			var text = sb.ToString();
			var to = $"{_username};{user};";

			Debug.WriteLine(text);

			await service.SendMailAsync(to, subject, text);
		}

		[TestMethod]
		[Owner(nameof(ExchangeOnlineManager))]
		[TestCategory("取得")]
		public void 会議室一覧を取得する() {
			var service = new ExchangeOnlineManager(_username, _password);

			var sb = new StringBuilder();

			foreach (var address in service.GetRooms()) {
				sb.AppendLine(address.GetPropertiesString());
			}

			Debug.WriteLine(sb.ToString());
		}

		[TestMethod]
		[Owner(nameof(ExchangeOnlineManager))]
		[TestCategory("取得")]
		public async Task 非同期で会議室一覧を取得する() {
			var service = new ExchangeOnlineManager(_username, _password);

			var sb = new StringBuilder();

			foreach (var address in await service.GetRoomsAsync()) {
				sb.AppendLine(address.GetPropertiesString());
			}

			Debug.WriteLine(sb.ToString());
		}

		[TestMethod]
		[Owner(nameof(ExchangeOnlineManager))]
		[TestCategory("作成")]
		public void 予定を作成する() {
			var mng = new ExchangeOnlineManager(_username, _password);

			var subject = "テニスレッスン";
			var body = "今週はバックハンドに焦点を当てる。";
			var start = new DateTime(2017, 5, 1, 10, 30, 0);
			var end = start.AddHours(1);
			var location = "テニスクラブ";
			var reminderDueBy = DateTime.Now;

			var appointment = mng.Save(subject, start, end, a => {
				a.Body = body;
				a.Location = location;
				a.ReminderDueBy = reminderDueBy;
			});

			// 予定のアイテムIDを使用して予定が作成されたことを確認します。
			var item = mng.Bind(appointment.Id.UniqueId, Ews.ItemSchema.Subject);

			Debug.WriteLine($@"Appointment created: {item.Subject}");
		}

		[TestMethod]
		[Owner(nameof(ExchangeOnlineManager))]
		[TestCategory("作成")]
		public async Task 非同期で予定を変更する() {
			var mng = new ExchangeOnlineManager(_username, _password);

			var aps = mng.FindAppointments(new DateTime(2017, 4, 1), new DateTime(2017, 6, 1));

			var ap = aps.FirstOrDefault(a => a.Subject.HasString("テニスレッスン"));

			var uniqueId = ap?.Id?.UniqueId;

			var appointment = await mng.UpdateAsync(uniqueId, a => {
				// 予定のプロパティを新しい件名、開始時刻、終了時刻で更新します。
				a.Subject = $"moved one hour later and to the day after {a.Start:yyyy/MM/dd(ddd)}!";
				a.Start = a.Start.AddHours(25);
				a.End = a.End.AddHours(25);
			});

			// 更新を確認します。
			Debug.WriteLine($@"""{appointment.Subject}""");
		}

		[TestMethod]
		[Owner(nameof(ExchangeOnlineManager))]
		[TestCategory("取得")]
		public async Task 非同期で会議室の予定を取得する() {
			var service = new ExchangeOnlineManager(_username, _password);

			// 出席者(会議室)のコレクションを作成します。 
			var attendees = await service.GetRoomsAsync();

			var start = DateTime.Today;
			var end = start.AddDays(7);

			var sc = new ExchangeScheduler(service, start, end, attendees) {
				GoodSuggestionThreshold = 49,
				MaximumNonWorkHoursSuggestionsPerDay = 8,
				MaximumSuggestionsPerDay = 8,
				MeetingDuration = 60,
				MinimumSuggestionQuality = Ews.SuggestionQuality.Good,
				RequestedFreeBusyView = Ews.FreeBusyViewType.FreeBusy,
			};

			var sb = new StringBuilder();
			var availabilities = await sc.GetUserAvailabilitiesAsync();

			availabilities.ForEach(info => {
				sb.AppendLine($"[{info.Key}]:");

				// 出席者のカレンダーイベントのコレクションを取得します。
				var days = info.Value.Select(ev => ev.StartTime.Date).Distinct();
				var schedule = info.Value;

				days.ForEach(day => {
					sb.AppendLine($"{day:yyyy/MM/dd(ddd)}");

					var blankPlans = day.GetBlankPlans();

					var dayPlans = schedule.Where(d => d.StartTime.Date == day);

					blankPlans.Where(p => !dayPlans.Any(d =>
						d.StartTime <= p.StartTime  && p.StartTime < d.EndTime
						|| d.StartTime < p.EndTime && p.EndTime <= d.EndTime
					)).ForEach(ev => {
						sb.AppendLine($"\t{ev.StartTime:HH:mm} ~ {ev.EndTime:HH:mm}");
					});
				});

				sb.AppendLine();
			});

			sb.AppendLine("--------------------------------------------------------------------------------");

			var subject = $"{sc.StartTime:yyyy/MM/dd(ddd)} ~ {sc.LastTime:yyyy/MM/dd(ddd)} の推奨される会議時間";
			var text = sb.ToString();
			var to = $"{_username}";

			Debug.WriteLine(text);

			await service.SendMailAsync(to, subject, text);
		}

		#endregion
	}

	public static partial class TimeExtension {
		public static IEnumerable<Ews.TimeWindow> GetBlankPlans(this DateTime @this) {
			var ret = @this.GetBlankPlansPerMinutes(9.0, 18.0, 30);
			var blankPlans = ret.ToCurrentAndNextPair()
				.Select(cn => new Ews.TimeWindow(cn.Item1, cn.Item2));
			return blankPlans;
		}

		public static LinkedList<DateTime> GetBlankPlansPerMinutes(this DateTime @this, double openingTime, double closingTime, int intervalPerMinutes) {
			var today = @this.Date;
			var b = new TimeSpan(Convert.ToInt64(TimeSpan.TicksPerHour * openingTime));
			var f = new TimeSpan(Convert.ToInt64(TimeSpan.TicksPerHour * closingTime));
			var ret = GetTimeSpans(b, f, intervalPerMinutes).Select(t => today + t).ToList();
			return new LinkedList<DateTime>(ret);
		}

		/// <summary>
		/// 現在の値と次の値のペアの列挙に変換します。
		/// </summary>
		/// <typeparam name="T"></typeparam>
		/// <param name="this"></param>
		/// <returns></returns>
		public static IEnumerable<Tuple<T, T>> ToCurrentAndNextPair<T>(this LinkedList<T> @this)
			=> @this
			.Where(node => node.Next != null)
			.Select(node => Tuple.Create(node.Value, node.Next.Value));

		private static IEnumerable<TimeSpan> GetTimeSpans(TimeSpan open, TimeSpan close, int intervalPerMinutes) {
			if (open > close) {
				throw new ArgumentException($"終了時刻が、開始時刻よりも前に設定されています。: {close}", nameof(close));
			}

			var tt = new TimeSpan(TimeSpan.TicksPerMinute * intervalPerMinutes);
			for (var m = open; m <= close; m += tt) {
				yield return m;
			}
		}

		private static IEnumerable<LinkedListNode<T>> ToEnumerable<T>(this LinkedList<T> @this) {
			for (var node = @this.First; node != null; node = node.Next) {
				yield return node;
			}
		}

		/// <summary>
		/// 各要素に対して、指定された処理を実行します。
		/// </summary>
		/// <typeparam name="T"></typeparam>
		/// <param name="this">LinkedList</param>
		/// <param name="action">各要素に対して実行するメソッド</param>
		public static void ForEach<T>(this LinkedList<T> @this, Action<LinkedListNode<T>> action) {
			@this.ToEnumerable().ToList().ForEach(action);
		}

		/// <summary>
		/// LinkedList の各要素を新しいフォームに射影します。
		/// </summary>
		/// <typeparam name="TSource">source の要素の型</typeparam>
		/// <typeparam name="TResult">selector によって返される値の型</typeparam>
		/// <param name="this">LinkedList</param>
		/// <param name="selector">各要素に適用する変換関数。</param>
		/// <returns>source の各要素に対して変換関数を呼び出した結果として得られる要素を含む列挙を返します。</returns>
		public static IEnumerable<TResult> Select<TSource, TResult>(this LinkedList<TSource> @this, Func<LinkedListNode<TSource>, TResult> selector)
			=> @this.ToEnumerable().Select(selector);

		/// <summary>
		/// 述語に基づいて値のシーケンスをフィルター処理します。
		/// </summary>
		/// <typeparam name="TSource">source の要素の型</typeparam>
		/// <param name="this">LinkedList</param>
		/// <param name="predicate">各要素が条件を満たしているかどうかをテストする関数。</param>
		/// <returns>条件を満たす、入力シーケンスの要素を含む列挙を返します。</returns>
		public static IEnumerable<LinkedListNode<TSource>> Where<TSource>(this LinkedList<TSource> @this, Func<LinkedListNode<TSource>, bool> predicate)
			=> @this.ToEnumerable().Where(predicate);
	}
}
