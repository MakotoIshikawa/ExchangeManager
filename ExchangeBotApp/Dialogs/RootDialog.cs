using System;
using System.Linq;
using System.Threading.Tasks;
using BotLibrary.Extensions;
using ExchangeBotApp.Extensions;
using ExchangeManager;
using ExchangeManager.Extensions;
using ExchangeManager.Model;
using ExtensionsLibrary.Extensions;
using JsonLibrary.Extensions;
using Microsoft.Bot.Builder.Dialogs;
using Microsoft.Bot.Connector;
using Ews = Microsoft.Exchange.WebServices.Data;

namespace ExchangeBotApp.Dialogs {
	/// <summary>
	/// 会議室情報を返信する Bot のダイアログです。
	/// </summary>
	[Serializable]
	public class RootDialog : IDialog<object> {
		#region フィールド

		private static string _username = null;
		private static string _password = null;

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

				var args = msg.Split(',', ';', '|', ' ');

				if (args.Length == 2 && args[0].IsMailAddress() && !args[1].IsEmpty()) {
					context.SetUserData(nameof(_username), args[0]);
					context.SetUserData(nameof(_password), args[1]);

					await context.PostAsync("ユーザ名とパスワードを設定しました。");
					return;
				}

				_username = context.GetUserData<string>(nameof(_username));
				_password = context.GetUserData<string>(nameof(_password));

				if (_username.IsEmpty() || _password.IsEmpty()) {
					throw new ApplicationException("ユーザ名とパスワードを入力して下さい。");
				}

				var manager = new ExchangeOnlineManager(_username, _password);

				if (msg.MatchWords("会議", "空")) {
					await context.PostAllScheduleAsync(manager);

					return;
				} else if (msg.MatchWords("address", "start", "end")) {
					var meeting = msg.Deserialize<MeetingModel>();

					await context.PostAsync($"{meeting.Location}を予約します。");

					var id = await manager.SaveAsync(meeting);
					var item = await manager.BindAsync(id);

					await context.PostAsync($"{item.Location}を予約しました。");

					return;
				} else if (msg.MatchWords("会議", "場所")) {
					var roomls = await manager.GetRoomListsAsync();
					var dic = roomls.ToDictionary(r => r.Name, r => r.Address);

					await context.PostButtonsAsync($"会議室のある場所の一覧です。", dic);

					return;
				} else if (msg.MatchWords("会議室")) {
					await context.PostListOfMeetingRoomsAsync(manager);

					return;
				}

				if (msg.IsMailAddress()) {
					//TODO: 会議室配布グループのアドレスなのか、会議室自体のアドレスなのかを判定する処理
					//TODO: 会議室配布グループのアドレスの場合、所属する会議室の一覧を表示する処理

					var rooms = await manager.GetRoomsAsync();
					var name = rooms.FirstOrDefault(r => r.Address == msg)?.Name;
					var address = new Ews.EmailAddress(name, msg);

					// 会議室の空き時間表示
					await context.PostConferenceScheduleAsync(manager, address);
					return;
				}

				await context.PostAsync($"「{msg}」ですか？");
			} catch (Exception ex) {
				await context.PostAsync(ex.Message);
			} finally {
				context.Wait(MessageReceivedAsync);
			}
		}

		#endregion
	}
}