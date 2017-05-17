using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Threading.Tasks;
using BotLibrary.Extensions;
using ExchangeManager;
using ExchangeManager.Interface;
using ExchangeManager.Model;
using ExtensionsLibrary.Extensions;
using Microsoft.Bot.Builder.Dialogs;
using Ews = Microsoft.Exchange.WebServices.Data;

namespace ExchangeBotApp.Extensions {
	/// <summary>
	/// DialogContext を拡張するメソッドを提供します。
	/// </summary>
	public static partial class DialogContextExtension {
		#region メソッド

		#region 会議室情報

		/// <summary>
		/// 全ての会議室の空席情報を返します。
		/// </summary>
		/// <param name="context">DialogContext</param>
		/// <param name="manager">Exchange を管理するオブジェクト</param>
		public static async Task PostAllScheduleAsync(this IDialogContext context, IExchangeManager manager) {
			var rooms = await manager.GetRoomsAsync();

			await context.PostAsync($"全ての会議室の空き状況をお調べします。");
			await PostConferenceScheduleAsync(context, manager, rooms.ToArray());
		}

		/// <summary>
		/// 会議室の一覧情報を返します。
		/// </summary>
		/// <param name="context">DialogContext</param>
		/// <param name="manager">Exchange を管理するオブジェクト</param>
		public static async Task PostListOfMeetingRoomsAsync(this IDialogContext context, IExchangeManager manager) {
			var rooms = await manager.GetRoomsAsync();
			var dic = rooms.ToDictionary(r => r.Name, r => r.Address);

			await context.PostButtonsAsync($"会議室の一覧です。", dic);
		}

		/// <summary>
		/// 会議室の空席情報を返します。
		/// </summary>
		/// <param name="context">DialogContext</param>
		/// <param name="manager">Exchange を管理するオブジェクト</param>
		/// <param name="addresses">メールアドレス</param>
		public static async Task PostConferenceScheduleAsync(this IDialogContext context, IExchangeManager manager, params Ews.EmailAddress[] addresses) {
			var now = DateTime.Now;
			var today = now.Date;
			var sc = new ExchangeScheduler(manager, today, addresses) {
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
					Attendees = new List<string> { manager.UserName, },
				}).ToDictionary(
					t => $"\t{t.Start:HH:mm} ~ {t.End:HH:mm}"
					, t => t.ToJson()
				);

				await context.PostButtonsAsync($"{mailBox.Name} : {today:yyyy/MM/dd(ddd)} 以下の時間帯が空いています。", dic);
			});
		}

		#endregion

		#endregion
	}
}