using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using ExchangeManager;
using ExchangeManager.Interface;
using ExchangeManager.Model;
using ExtensionsLibrary.Extensions;
using Microsoft.Bot.Builder.Dialogs;
using Microsoft.Bot.Builder.Luis.Models;
using Microsoft.Bot.Connector;
using Ews = Microsoft.Exchange.WebServices.Data;

namespace ExchangeBotApp.Extensions {
	/// <summary>
	/// DialogContext を拡張するメソッドを提供します。
	/// </summary>
	public static partial class DialogContextExtension {
		#region メソッド

		#region ユーザーデータ

		public static T GetUserData<T>(this IDialogContext @this, string key)
			=> @this.UserData.GetValueOrDefault(u => u.GetValue<T>(key));

		/// <summary>
		/// ユーザーデータに値を設定します。
		/// </summary>
		/// <typeparam name="T">値の型</typeparam>
		/// <param name="this">DialogContext</param>
		/// <param name="key">キー</param>
		/// <param name="value">値</param>
		public static void SetUserData<T>(this IDialogContext @this, string key, T value)
			=> @this.UserData.SetValue(key, value);

		#endregion

		#region PostDefaultResponseMessageAsync

		/// <summary>
		/// デフォルト時の応答メッセージを POST します。
		/// </summary>
		/// <param name="this">IDialogContext</param>
		/// <param name="result">LUIS の結果</param>
		public static async Task PostDefaultResponseMessageAsync(this IDialogContext @this, LuisResult result, Dictionary<string, string> dictionary, Exception ex = null) {
			var msg = GetDefaultResponseMessage(result, ex);
			await @this.PostButtonsAsync(msg, dictionary);
		}

		/// <summary>
		/// デフォルト時の応答メッセージを取得します。
		/// </summary>
		/// <param name="result">LUIS の結果</param>
		/// <returns>応答メッセージを返します。</returns>
		private static string GetDefaultResponseMessage(this LuisResult result, Exception ex = null) {
			var sb = new StringBuilder();
			sb.AppendLine("お問い合わせありがとうございます。");

			if (ex != null) {
				sb.AppendLine($" {ex.Message} ");
			}

			sb.AppendLine("恐れ入りますが以下のサイトでお調べください。")
			.AppendLine($@">https://www.google.co.jp/search?q={result.Query}");

			return sb.ToString();
		}

		#endregion

		#region PostButtonsAsync

		public static async Task PostButtonsAsync(this IDialogContext context, string text, Dictionary<string, string> dictionary) {
			var btns = dictionary?.ToButtons();
			await context.PostMessageAsync(text, btns);
		}

		private static Attachment ToButtons(this Dictionary<string, string> dictionary, string text = null) {
			var cd = new ThumbnailCard {
				Buttons = (
					from kv in dictionary
					select kv.ToButton()
				).ToList(),
			};
			return cd.ToAttachment();
		}

		private static CardAction ToButton(this KeyValuePair<string, string> kv) {
			return new CardAction() {
				Value = kv.Value,
				Type = "postBack",
				Title = kv.Key,
			};
		}

		#endregion

		#region PostMessageAsync

		/// <summary>
		/// 回答を非同期で POST します。
		/// </summary>
		/// <param name="context">IDialogContext</param>
		/// <param name="text">テキスト文字列</param>
		/// <param name="attachments">付属品の配列</param>
		/// <returns></returns>
		public static async Task PostMessageAsync(this IDialogContext context, string text, params Attachment[] attachments) {
			var msg = context.CreateMessage();

			if (!text.IsEmpty()) {
				msg.Text = text;
			}

			if (attachments?.Any(a => a != null) ?? false) {
				attachments.ForEach(a => msg.Attachments.Add(a));
			}

			await context.PostAsync(msg);
		}

		private static IMessageActivity CreateMessage(this IDialogContext context) {
			var msg = context.MakeMessage();
			if (msg.Attachments == null) {
				msg.Attachments = new List<Attachment>();
			}

			return msg;
		}

		#endregion

		#region SendAsync

		/// <summary>
		/// 非同期でダイアログの処理を開始します。
		/// </summary>
		/// <typeparam name="TDialog">Dialog</typeparam>
		/// <param name="this">Activity</param>
		public static async Task SendAsync<TDialog>(this Activity @this)
			where TDialog : IDialog<object>, new() {
			await Conversation.SendAsync(@this, () => new TDialog());
		}

		#endregion

		#region PostAsync

		/// <summary>
		/// 非同期で POST します。
		/// </summary>
		/// <typeparam name="TDialog">ダイアログの型</typeparam>
		/// <param name="activity">Activity</param>
		/// <param name="handleSystemMessage">システムメッセージを処理するメソッド</param>
		/// <returns>HTTP 応答メッセージを返します。</returns>
		public static async Task<HttpResponseMessage> PostAsync<TDialog>(this Activity activity, Action<Activity> handleSystemMessage = null) where TDialog : IDialog<object>, new() {
			var type = activity?.GetActivityType();
			if (type == ActivityTypes.Message) {
				await activity.SendAsync<TDialog>();
			} else {
				handleSystemMessage?.Invoke(activity);
			}

			var response = new HttpResponseMessage(HttpStatusCode.Accepted);
			return response;
		}

		#endregion

		/// <summary>
		/// 全ての会議室の空席情報を返します。
		/// </summary>
		/// <param name="context">DialogContext</param>
		public static async Task PostAllScheduleAsync(this IDialogContext context, IExchangeManager manager) {
			var rooms = await manager.GetRoomsAsync();

			await context.PostAsync($"全ての会議室の空き状況をお調べします。");
			await PostConferenceScheduleAsync(context, manager, rooms.ToArray());
		}

		/// <summary>
		/// 会議室の一覧情報を返します。
		/// </summary>
		/// <param name="context">DialogContext</param>
		public static async Task PostListOfMeetingRoomsAsync(this IDialogContext context, IExchangeManager manager) {
			var rooms = await manager.GetRoomsAsync();
			var dic = rooms.ToDictionary(r => r.Name, r => r.Address);

			await context.PostButtonsAsync($"会議室の一覧です。", dic);
		}

		/// <summary>
		/// 会議室の空席情報を返します。
		/// </summary>
		/// <param name="context">DialogContext</param>
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

		#region MatchWords

		/// <summary>
		/// 指定したワードの配列を全て含んでいるかどうか判定します。
		/// </summary>
		/// <param name="this">string</param>
		/// <param name="words">ワードの配列</param>
		/// <returns>全てのワードを含んでいる場合は、true を返します。</returns>
		public static bool MatchWords(this string @this, params string[] words)
			=> words.All(s => @this.HasString(s));

		#endregion

		#endregion
	}
}