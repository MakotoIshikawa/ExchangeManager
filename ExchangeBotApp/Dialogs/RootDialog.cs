using System;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExchangeBotApp.Extensions;
using ExchangeBotApp.Properties;
using ExchangeManager;
using ExchangeManager.Extensions;
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

				if (IsMailAddress(msg)) {
					var service = new ExchangeOnlineManager(_username, _password);
					var now = DateTime.Now;
					var today = now.Date;
					var sc = new ExchangeScheduler(service, today, new Ews.AttendeeInfo(msg)) {
						GoodSuggestionThreshold = 49,
						MaximumNonWorkHoursSuggestionsPerDay = 8,
						MaximumSuggestionsPerDay = 8,
						MeetingDuration = 60,
						MinimumSuggestionQuality = Ews.SuggestionQuality.Good,
						RequestedFreeBusyView = Ews.FreeBusyViewType.FreeBusy,
					};

					var openingTime = 9.0;
					var closingTime = 18.0;
					var intervalPerMinutes = 30;

					var sb = new StringBuilder();
					sb.AppendLine($"{today:yyyy/MM/dd(ddd)} 以下の時間帯が空いています。");

					var availabilities = await sc.GetUserAvailabilitiesAsync();

					var info = availabilities.First();
					var times = info.Value.GetBlankTimes(openingTime, closingTime, intervalPerMinutes)
						.Where(t => t.StartTime > now);

					var dic = times.ToDictionary(t => $"\t{t.StartTime:HH:mm} ~ {t.EndTime:HH:mm}", t => t.GetPropertiesString());

					await context.PostButtonsAsync(sb.ToString(), dic);

					return;
				}

				if (msg.MatchWords("会議", "空")) {
					await context.PostAsync($"会議室の空き時間です。");

					return;
				} else if (msg.MatchWords("会議", "場所")) {
					var service = new ExchangeOnlineManager(_username, _password);
					var roomls = await service.GetRoomListsAsync();
					var dic = roomls.ToDictionary(r => r.Name, r => r.Address);

					await context.PostButtonsAsync($"会議室の場所です。", dic);

					return;
				} else if (msg.MatchWords("会議室")) {
					var service = new ExchangeOnlineManager(_username, _password);
					var rooms = await service.GetRoomsAsync();
					var dic = rooms.ToDictionary(r => r.Name, r => r.Address);

					await context.PostButtonsAsync($"会議室の一覧です。", dic);

					return;
				}

				await context.PostAsync($"「{msg}」ですか？");
			} catch (Exception ex) {
				await context.PostAsync(ex.Message);
			} finally {
				context.Wait(MessageReceivedAsync);
			}
		}

		private static bool IsMailAddress(string address) {
			try {
				var a = new System.Net.Mail.MailAddress(address);

				return true;
			} catch (Exception) {
				return false;
			}
		}

		#endregion
	}
}