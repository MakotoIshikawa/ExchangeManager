using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using ExchangeBotApp.Extensions;
using ExchangeBotApp.Properties;
using ExchangeManager;
using ExchangeManager.Extensions;
using ExchangeManager.Model;
using ExtensionsLibrary.Extensions;
using Microsoft.Bot.Builder.Dialogs;
using Microsoft.Bot.Connector;
using Ews = Microsoft.Exchange.WebServices.Data;

namespace ExchangeBotApp.Dialogs {
	[Serializable]
	public class RootDialog : IDialog<object> {
		#region フィールド

		private static string _username = Settings.Default.UserName;
		private static string _password = Settings.Default.Password;
		private static ExchangeOnlineManager _manager = new ExchangeOnlineManager(_username, _password);

		#endregion

		#region メソッド

		#region IDialog

		/// <summary>
		/// 会話ダイアログを表すコードの開始点
		/// </summary>
		/// <param name="context">ダイアログコンテキスト</param>
		public async Task StartAsync(IDialogContext context)
			=> await Task.Run(() => context.Wait(this.MessageReceivedAsync));

		#endregion

		private async Task MessageReceivedAsync(IDialogContext context, IAwaitable<IMessageActivity> argument) {
			try {
				var activity = await argument as Activity;
				var request = activity.Text.Trim();

				var msg = request.TrimEnd("、。,.".ToArray());
				if (msg.IsEmpty()) {
					throw new ApplicationException("何かメッセージを入れてください。");
				}

				if (msg.MatchWords("会議", "空")) {
					await PostAllScheduleAsync(context);

					return;
				} else if (msg.MatchWords("address", "start", "end")) {
					var meeting = msg.Deserialize<MeetingModel>();

					await context.PostAsync($"{meeting.Location}を予約します。");

					var id = await _manager.SaveAsync(meeting);

					var item = await _manager.BindAsync(id);

					await context.PostAsync($"{item.Location}を予約しました。");

					return;
				} else if (msg.MatchWords("会議", "場所")) {
					var roomls = await _manager.GetRoomListsAsync();
					var dic = roomls.ToDictionary(r => r.Name, r => r.Address);

					await context.PostButtonsAsync($"会議室のある場所の一覧です。", dic);

					return;
				} else if (msg.MatchWords("会議室")) {
					await PostListOfMeetingRoomsAsync(context);

					return;
				}

				if (msg.IsMailAddress()) {
					await PostConferenceScheduleAsync(context, msg);
					return;
				}

				await context.PostAsync($"「{msg}」ですか？");
			} catch (Exception ex) {
				await context.PostAsync(ex.Message);
			} finally {
				context.Wait(MessageReceivedAsync);
			}
		}

		/// <summary>
		/// 全ての会議室の空席情報を返します。
		/// </summary>
		/// <param name="context">DialogContext</param>
		private async Task PostAllScheduleAsync(IDialogContext context) {
			var rooms = await _manager.GetRoomsAsync();

			await context.PostAsync($"全ての会議室の空き状況をお調べします。");
			await PostConferenceScheduleAsync(context, rooms.ToArray());
		}

		/// <summary>
		/// 会議室の一覧情報を返します。
		/// </summary>
		/// <param name="context">DialogContext</param>
		private async Task PostListOfMeetingRoomsAsync(IDialogContext context) {
			var rooms = await _manager.GetRoomsAsync();
			var dic = rooms.ToDictionary(r => r.Name, r => r.Address);

			await context.PostButtonsAsync($"会議室の一覧です。", dic);
		}

		/// <summary>
		/// 会議室の空席情報を返します。
		/// </summary>
		/// <param name="context">DialogContext</param>
		/// <param name="addresses">メールアドレス</param>
		private async Task PostConferenceScheduleAsync(IDialogContext context, params Ews.EmailAddress[] addresses) {
			var now = DateTime.Now;
			var today = now.Date;
			var sc = new ExchangeScheduler(_manager, today, addresses) {
				GoodSuggestionThreshold = 49,
				MaximumNonWorkHoursSuggestionsPerDay = 8,
				MaximumSuggestionsPerDay = 8,
				MeetingDuration = 60,
				MinimumSuggestionQuality = Ews.SuggestionQuality.Good,
				RequestedFreeBusyView = Ews.FreeBusyViewType.FreeBusy,
				OpeningTime = 9.0,
				ClosingTime = 18.0,
				IntervalPerMinutes = 30,
			};

			var times = await sc.GetBlankTimesAsync();

			if (!(times?.SelectMany(t => t.Item2)?.Any(t => t.StartTime > now) ?? false)) {
				await context.PostAsync($"{today:yyyy/MM/dd(ddd)} 現在、空いてる会議室がありません。");
				return;
			}

			times.ForEach(async time => {
				var mailBox = time.Item1;
				var ts = time.Item2.Where(t => t.StartTime > now);
				if (!ts.Any()) {
					await context.PostAsync($"{mailBox.Name} : {today:yyyy/MM/dd(ddd)} 空いてる時間帯はありません。");
					return;
				}

				var dic = ts.Select(t => new MeetingModel("会議", mailBox, t.StartTime, t.EndTime) {
					Body = "Bot から予約された会議です。",
					Attendees = new List<string> { _username },
				}).ToDictionary(
					t => $"\t{t.Start:HH:mm} ~ {t.End:HH:mm}"
					, t => t.ToJson()
				);

				await context.PostButtonsAsync($"{mailBox.Name} : {today:yyyy/MM/dd(ddd)} 以下の時間帯が空いています。", dic);
			});
		}

		#endregion
	}
}